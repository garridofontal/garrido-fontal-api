const express = require('express');
const cors    = require('cors');
const path    = require('path');
const multer  = require('multer');
const { Packer } = require('docx');
const { buildPresuposto }   = require('./builders/presuposto');
const { buildFactura }      = require('./builders/factura');
const { readPresuposto }    = require('./builders/readPresuposto');

const app    = express();
const PORT   = process.env.PORT || 3000;
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });

app.use(cors());
app.use(express.json({ limit: '2mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ── Health check ─────────────────────────────────────────────────────────────
app.get('/health', (_req, res) => res.json({ ok: true }));

// ── Ler presuposto (extract data from uploaded .docx) ─────────────────────────
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

// ── Generar presupuesto ───────────────────────────────────────────────────────
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

// ── Generar factura ───────────────────────────────────────────────────────────
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
