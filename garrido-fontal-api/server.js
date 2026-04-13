const express = require('express');
const cors    = require('cors');
const path    = require('path');
const { Packer } = require('docx');
const { buildPresuposto } = require('./builders/presuposto');
const { buildFactura }    = require('./builders/factura');

const app  = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json({ limit: '2mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// ── Health check ─────────────────────────────────────────────────────────────
app.get('/health', (_req, res) => res.json({ ok: true }));

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
