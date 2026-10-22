require('dotenv').config({ path: require('path').join(__dirname, '.env') });
const express = require('express');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { spawnSync } = require('child_process');
const Anthropic = require('@anthropic-ai/sdk');

const app = express();
app.use(express.json({ limit: '50mb' }));
app.use(express.static(__dirname));

const hasApiKey = !!process.env.ANTHROPIC_API_KEY && process.env.ANTHROPIC_API_KEY !== 'your_anthropic_api_key_here';
if (!hasApiKey) {
  console.warn('\n⚠️   ANTHROPIC_API_KEY not set — AI detection disabled. Offline mode only.\n');
}

const client = hasApiKey ? new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY }) : null;
const SYSTEM_PROMPT = fs.readFileSync(path.join(__dirname, 'system-prompt.txt'), 'utf8');
const CORRECTIONS_FILE = path.join(__dirname, 'corrections.json');

// ─── Corrections helpers ──────────────────────────────────────────────────────

function loadCorrections() {
  try {
    if (fs.existsSync(CORRECTIONS_FILE)) {
      return JSON.parse(fs.readFileSync(CORRECTIONS_FILE, 'utf8'));
    }
  } catch (e) {
    console.warn('Could not load corrections:', e.message);
  }
  return [];
}

function saveCorrections(list) {
  try {
    fs.writeFileSync(CORRECTIONS_FILE, JSON.stringify(list, null, 2));
  } catch (e) {
    // Vercel has a read-only filesystem — corrections are in-memory only
    console.warn('Could not persist corrections (read-only fs):', e.message);
  }
}

function buildFewShotBlock(corrections) {
  if (!corrections.length) return '';
  const recent = corrections.slice(-8); // last 8 corrections as context
  const lines = recent.map((c, i) => {
    if (c.type === 'injection_failure') {
      return `${i + 1}. [INJECTION FAILURE] Field "${c.fieldName}" — error: ${c.error}. Target was: "${c.target}". AnchorText: "${c.anchorText}".`;
    }
    if (c.type === 'field_rename') {
      return `${i + 1}. [FIELD RENAME] "${c.originalName}" corrected to "${c.correctedName}" in context: "${c.context}".`;
    }
    if (c.type === 'field_reject') {
      return `${i + 1}. [FIELD REJECTED] Field "${c.fieldName}" was rejected — do not create this field for context: "${c.context}".`;
    }
    return `${i + 1}. [CORRECTION] ${JSON.stringify(c)}`;
  });
  return `<previous-corrections>\nUse these past corrections to improve your field detection accuracy:\n${lines.join('\n')}\n</previous-corrections>\n\n`;
}

// ─── DOCX → PDF via Word COM (Windows only — not available on Vercel) ────────

app.post('/api/convert-to-pdf', (req, res) => {
  if (process.platform !== 'win32') {
    return res.status(501).json({ error: 'DOCX→PDF conversion is only available when running locally on Windows.' });
  }
  const { docxBase64, filename } = req.body;
  if (!docxBase64) return res.status(400).json({ error: 'docxBase64 required' });

  const safeName = (filename || 'doc').replace(/[^a-zA-Z0-9._-]/g, '_');
  const stamp = Date.now();
  const docxPath = path.join(os.tmpdir(), `lfi_${stamp}_${safeName}`);
  const pdfPath  = docxPath.replace(/\.docx$/i, '') + '.pdf';
  const psPath   = path.join(os.tmpdir(), `lfi_${stamp}.ps1`);

  try {
    fs.writeFileSync(docxPath, Buffer.from(docxBase64, 'base64'));

    // Use Word COM to convert DOCX → PDF via ExportAsFixedFormat
    // $word.Quit() has no required params — avoid passing 0 which causes [ref] errors
    const dPath = docxPath.replace(/\\/g, '\\\\');
    const pPath = pdfPath.replace(/\\/g, '\\\\');
    const psScript = `
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  $d = $word.Documents.Open([ref]'${dPath}', [ref]$false, [ref]$true, [ref]$false)
  $d.ExportAsFixedFormat('${pPath}', 17)
  $d.Close()
} finally {
  try { $word.Quit() } catch {}
}
`;
    fs.writeFileSync(psPath, psScript, 'utf8');

    // Use spawnSync so a non-zero exit from Word COM cleanup doesn't mask success
    const result = spawnSync(
      'powershell',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', psPath],
      { timeout: 120000 }
    );

    // Success = PDF file was written, regardless of PS exit code (Word COM cleanup may exit 1)
    if (!fs.existsSync(pdfPath)) {
      const stderr = result.stderr ? result.stderr.toString().trim() : '';
      throw new Error(`Word COM conversion failed: ${stderr || 'PDF not created'}`);
    }

    const pdfBuffer = fs.readFileSync(pdfPath);
    res.json({ pdfBase64: pdfBuffer.toString('base64') });
    console.log(`✓ DOCX→PDF converted: ${safeName} (${Math.round(pdfBuffer.length / 1024)}KB)`);
  } catch (err) {
    console.error('convert-to-pdf error:', err.message);
    res.status(500).json({ error: err.message });
  } finally {
    [docxPath, pdfPath, psPath].forEach(f => { try { fs.unlinkSync(f); } catch {} });
  }
});

// ─── Per-page field detection (for canvas-rendered / image-based documents) ───

app.post('/api/detect-page', async (req, res) => {
  if (!client) return res.status(503).json({ error: 'ANTHROPIC_API_KEY not configured.' });
  const {
    imageBase64, mediaType, pageIndex, pageCount,
    documentType, extraInstructions, pageSpecificInstructions,
    continuityContext, model: modelOverride,
  } = req.body;
  if (!imageBase64) return res.status(400).json({ error: 'imageBase64 required' });

  const model = modelOverride || 'claude-sonnet-4-5';

  try {
    const corrections = loadCorrections();
    const fewShot = buildFewShotBlock(corrections);

    const contextParts = [];
    if (documentType) contextParts.push(`Document type: ${documentType}`);
    if (extraInstructions) contextParts.push(`Extra instructions: ${extraInstructions}`);
    if (pageSpecificInstructions) contextParts.push(`Page ${(pageIndex || 0) + 1} specific instructions: ${pageSpecificInstructions}`);
    if (continuityContext && continuityContext.length) {
      contextParts.push(`Fields already detected on previous pages — use consistent naming: ${continuityContext.slice(-30).join(', ')}`);
    }

    const pagePrompt = [
      fewShot || '',
      contextParts.length ? `<document-context>\n${contextParts.join('\n')}\n</document-context>` : '',
      `You are analyzing page ${(pageIndex || 0) + 1} of ${pageCount || 1} of a bank/government form.`,
      `Return ONLY a JSON array of ALL fillable fields visible on this page.`,
      `CRITICAL: Use method "overlay" for every field. Set pageIndex: ${pageIndex || 0} on every field.`,
      ``,
      SYSTEM_PROMPT,
    ].filter(Boolean).join('\n');

    const response = await client.messages.create({
      model,
      max_tokens: 16000,
      system: pagePrompt,
      messages: [{
        role: 'user',
        content: [
          {
            type: 'text',
            text: `Detect ALL fillable fields on this page (page ${(pageIndex || 0) + 1} of ${pageCount || 1}). ` +
                  `Return JSON array. Every field must use method:"overlay" and pageIndex:${pageIndex || 0}. ` +
                  `Include every text box group, checkbox, radio option, and blank line on this page.`,
          },
          {
            type: 'image',
            source: { type: 'base64', media_type: mediaType || 'image/png', data: imageBase64 },
          },
        ],
      }],
    });

    let raw = response.content[0].text.trim()
      .replace(/^```(?:json)?\s*/i, '').replace(/\s*```\s*$/, '');

    let fields = [];
    try {
      const fullMatch = raw.match(/\[[\s\S]*\]/);
      fields = JSON.parse(fullMatch ? fullMatch[0] : raw);
    } catch {
      const partial = (raw.match(/\[[\s\S]*/) || ['[]'])[0];
      const lastCommaObj = partial.lastIndexOf('},');
      const lastObj = partial.lastIndexOf('}');
      const cut = lastCommaObj > lastObj - 5 ? lastCommaObj + 1 : lastObj + 1;
      try { fields = JSON.parse(partial.slice(0, cut) + ']'); } catch { fields = []; }
    }

    // Ensure correct pageIndex and method on every returned field
    const pi = pageIndex || 0;
    fields = fields.map(f => ({
      ...f,
      injectionPoint: {
        ...f.injectionPoint,
        method: 'overlay',
        pageIndex: pi,
      },
    }));

    console.log(`   Page ${pi + 1}/${pageCount}: ${fields.length} fields [${model}] (${Math.round(raw.length / 1024)}KB response)`);
    res.json({ fields, usage: response.usage });
  } catch (err) {
    console.error(`detect-page[${pageIndex}] error:`, err.message);
    res.status(500).json({ error: err.message });
  }
});

// ─── Routes ───────────────────────────────────────────────────────────────────

// Extract base64 images from HTML and replace with placeholders.
// Returns { strippedHtml, images: [{mediaType, data}] }
function extractBase64Images(html) {
  const images = [];
  const strippedHtml = html.replace(
    /src="data:(image\/(?:png|jpeg|jpg|gif|webp));base64,([A-Za-z0-9+/=]+)"/g,
    (match, mediaType, data) => {
      const idx = images.length;
      images.push({ mediaType, data });
      return `src="[IMAGE:${idx}]"`;
    }
  );
  return { strippedHtml, images };
}

// Detect fields via Claude
app.post('/api/detect-fields', async (req, res) => {
  if (!client) return res.status(503).json({ error: 'ANTHROPIC_API_KEY not configured on server. Using offline detection.' });
  const { html } = req.body;
  if (!html) return res.status(400).json({ error: 'html is required' });

  try {
    const corrections = loadCorrections();
    const fewShot = buildFewShotBlock(corrections);

    const { strippedHtml, images } = extractBase64Images(html);
    const docText = fewShot
      ? `${fewShot}---\n\nDocument HTML to analyze:\n\n${strippedHtml}`
      : strippedHtml;

    // Build multipart message: text first, then images as vision blocks
    const contentBlocks = [{ type: 'text', text: docText }];
    images.forEach((img, i) => {
      contentBlocks.push({ type: 'text', text: `\n[IMAGE:${i}] — Page ${i + 1}:` });
      contentBlocks.push({ type: 'image', source: { type: 'base64', media_type: img.mediaType, data: img.data } });
    });

    console.log(`→ Sending to Claude (${Math.round(docText.length / 1024)}KB text + ${images.length} image(s))…`);

    const response = await client.messages.create({
      model: 'claude-sonnet-4-5',
      max_tokens: 16000,
      system: SYSTEM_PROMPT,
      messages: [{ role: 'user', content: contentBlocks }],
    });

    // Strip markdown code fences if Claude wraps the JSON in ```json ... ```
    const raw = response.content[0].text.trim()
      .replace(/^```(?:json)?\s*/i, '')
      .replace(/\s*```\s*$/, '');
    console.log(`← Received response (${Math.round(raw.length / 1024)}KB)`);

    // Extract JSON array — handle truncated responses (token limit mid-array)
    const match = raw.match(/\[[\s\S]*/);
    if (!match) throw new Error('Claude did not return a JSON array. Raw: ' + raw.slice(0, 200));

    let fields;
    try {
      // Try full parse first (response ends with ] normally)
      const fullMatch = raw.match(/\[[\s\S]*\]/);
      fields = JSON.parse(fullMatch ? fullMatch[0] : match[0]);
    } catch (e) {
      // Response was truncated — recover by closing the last complete object
      let partial = match[0];
      // Remove everything after the last complete "}," or "}" before truncation
      const lastObj = partial.lastIndexOf('},');
      const lastObjFinal = partial.lastIndexOf('}');
      const cutAt = lastObj > lastObjFinal - 5 ? lastObj + 1 : lastObjFinal + 1;
      try {
        fields = JSON.parse(partial.slice(0, cutAt) + ']');
        console.warn(`⚠️  Response was truncated — recovered ${fields.length} complete fields`);
      } catch (e2) {
        throw new Error('Response JSON was malformed or too truncated to recover. Raw: ' + raw.slice(0, 300));
      }
    }
    console.log(`   Detected ${fields.length} fields`);

    res.json({ fields, usage: response.usage });
  } catch (err) {
    console.error('detect-fields error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// Save a correction for future few-shot learning
app.post('/api/save-correction', (req, res) => {
  try {
    const corrections = loadCorrections();
    corrections.push({ ...req.body, timestamp: Date.now() });
    // Keep last 50 corrections
    const trimmed = corrections.slice(-50);
    saveCorrections(trimmed);
    console.log(`✓ Correction saved (${trimmed.length} total)`);
    res.json({ ok: true, total: trimmed.length });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Get all corrections (for debug/admin)
app.get('/api/corrections', (req, res) => {
  res.json(loadCorrections());
});

// Delete a correction by index
app.delete('/api/corrections/:index', (req, res) => {
  try {
    const corrections = loadCorrections();
    const idx = parseInt(req.params.index);
    if (idx < 0 || idx >= corrections.length) return res.status(404).json({ error: 'Not found' });
    corrections.splice(idx, 1);
    saveCorrections(corrections);
    res.json({ ok: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── Start ────────────────────────────────────────────────────────────────────

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`\n✅  Doc Form Filler  →  http://localhost:${PORT}\n`);
});
