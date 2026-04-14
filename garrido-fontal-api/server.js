const express = require('express');
const cors = require('cors');
const path = require('path');
const multer = require('multer');
const { Packer } = require('docx');
const { buildPresuposto } = require('./builders/presuposto');
const { buildFactura } = require('./builders/factura');
const { readPresuposto } = require('./builders/readPresuposto');

const app = express();
const PORT = process.env.PORT || 3000;
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });

app.use(cors());
app.use(express.json({ limit: '2mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ── Health check ────────────────────────────────────────────────────────────────
app.get('/health', (_req, res) => res.json({ ok: true }));

// ── Traducir gallego → castellano ───────────────────────────────────────────────
app.post('/api/traducir', async (req, res) => {
  try {
    const { texto } = req.body;
    if (!texto || !texto.trim()) return res.json({ traduccion: '' });

    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) return res.status(500).json({ error: 'ANTHROPIC_API_KEY non configurada' });

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json',
      },
      body: JSON.stringify({
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 512,
        messages: [{
          role: 'user',
          content: `Traduce este texto do galego ó castelán. Devolve UNICAMENTE a tradución, sen explicacións, sen comillas, sen prefixos.\n\nTexto:\n${texto}`,
        }],
      }),
    });

    if (!response.ok) {
      const err = await response.text();
      return res.status(500).json({ error: 'API error: ' + err });
    }

    const data = await response.json();
    const traduccion = (data.content?.[0]?.text || '').trim();
    res.json({ traduccion });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// ── Ler presuposto (extract data from uploaded .docx) ───────────────────────────
app.post('/api/ler-presuposto', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No se recibiu ningún arquivo' });
    const mammoth = require('mammoth');
    const raw = await mammoth.extractRawText({ buffer: req.file.buffer });
    const data = await readPresuposto(req.file.buffer);
    data.rawLines = raw.value.split('\n').map(l => l.trim()).filter(Boolean);
    res.json(data);
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// ── Generar presupuesto ─────────────────────────────────────────────────────────
app.post('/api/presuposto', async (req, res) => {
  try {
    const doc = buildPresuposto(req.body);
    const buf = await Packer.toBuffer(doc);
    const filename = `Orzamento_${(req.body.num || 'draft').replace(/\//g, '-')}.docx`;
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.send(buf);
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// ── Generar factura ─────────────────────────────────────────────────────────────
app.post('/api/factura', async (req, res) => {
  try {
    const doc = buildFactura(req.body);
    const buf = await Packer.toBuffer(doc);
    const filename = `Factura_${(req.body.num || 'draft').replace(/\//g, '-')}.docx`;
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.send(buf);
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

app.listen(PORT, () => console.log(`Garrido Fontal API → http://localhost:${PORT}`));
