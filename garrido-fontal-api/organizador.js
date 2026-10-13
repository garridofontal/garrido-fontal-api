// ── API del Organizador de Trabajos ─────────────────────────────────────────
// Todas las rutas van bajo /api/org y (salvo /login) piden token.
const express = require('express');
const crypto  = require('crypto');
const fs      = require('fs');
const path    = require('path');
const { q, pool } = require('./db');
const calendario = require('./calendario');
const multer  = require('multer');
const { Packer } = require('docx');
const { buildPresuposto } = require('./builders/presuposto');
const { buildFactura }    = require('./builders/factura');
const { readPresuposto }  = require('./builders/readPresuposto');
const subida = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024, files: 20 } });

const router = express.Router();
const ESTADOS = ['Contacto', 'Pendiente de ver', 'Pendiente de presupuesto', 'Pendiente de hacer', 'Pendiente de cobrar', 'Terminado'];

// ── Autenticación por PIN ───────────────────────────────────────────────────
function secret() {
  return process.env.SESSION_SECRET ||
    crypto.createHash('sha256').update('gf:' + (process.env.APP_PIN || '') + ':' + (process.env.DATABASE_URL || '')).digest('hex');
}
function sign(payload) {
  const body = Buffer.from(JSON.stringify(payload)).toString('base64url');
  const mac  = crypto.createHmac('sha256', secret()).update(body).digest('base64url');
  return body + '.' + mac;
}
function verify(token) {
  if (!token || !token.includes('.')) return null;
  const [body, mac] = token.split('.');
  const ok = crypto.createHmac('sha256', secret()).update(body).digest('base64url');
  if (mac.length !== ok.length || !crypto.timingSafeEqual(Buffer.from(mac), Buffer.from(ok))) return null;
  const p = JSON.parse(Buffer.from(body, 'base64url').toString());
  return p.exp > Date.now() ? p : null;
}
const intentos = new Map();
router.post('/login', (req, res) => {
  if (!process.env.APP_PIN) return res.status(503).json({ error: 'Falta configurar APP_PIN en Railway' });
  const ip = req.headers['x-forwarded-for'] || req.ip;
  const i = intentos.get(ip) || { n: 0, t: 0 };
  if (i.n >= 5 && Date.now() - i.t < 10 * 60 * 1000) return res.status(429).json({ error: 'Demasiados intentos. Espera 10 minutos.' });
  const pin = String((req.body || {}).pin || '');
  const a = crypto.createHash('sha256').update(pin).digest();
  const b = crypto.createHash('sha256').update(process.env.APP_PIN).digest();
  if (!crypto.timingSafeEqual(a, b)) { intentos.set(ip, { n: i.n + 1, t: Date.now() }); return res.status(401).json({ error: 'PIN incorrecto' }); }
  intentos.delete(ip);
  res.json({ token: sign({ u: 'jose', exp: Date.now() + 90 * 24 * 3600 * 1000 }) });
});
router.use((req, res, next) => {
  if (!pool) return res.status(503).json({ error: 'Base de datos non configurada' });
  const t = (req.headers.authorization || '').replace(/^Bearer /, '') || req.query.token;
  if (!verify(t)) return res.status(401).json({ error: 'Sesión caducada' });
  next();
});

const wrap = (fn) => (req, res) => fn(req, res).catch((e) => { console.error(e); res.status(500).json({ error: e.message }); });
const num = (v) => (v === null || v === undefined ? 0 : parseFloat(v));

// ── Resumen para la pantalla de inicio ──────────────────────────────────────
router.get('/resumen', wrap(async (_req, res) => {
  const c = await q(`SELECT estado, count(*)::int n FROM trabajos WHERE cancelado_en IS NULL GROUP BY estado`);
  const cancel = await q(`SELECT count(*)::int n FROM trabajos WHERE cancelado_en IS NOT NULL`);
  const cobrar = await q(`SELECT coalesce(sum(d.base),0) s FROM documentos d JOIN trabajos t ON t.id=d.trabajo_id
                          WHERE d.tipo='factura' AND t.estado=4 AND t.cancelado_en IS NULL`);
  const porEstado = ESTADOS.map(() => 0);
  c.rows.forEach((r) => { porEstado[r.estado] = r.n; });
  res.json({ estados: ESTADOS, porEstado, cancelados: cancel.rows[0].n, porCobrar: num(cobrar.rows[0].s) });
}));

// ── Trabajos ────────────────────────────────────────────────────────────────
const TRABAJO_SELECT = `
  SELECT t.*, c.nombre cliente, c.tipo cliente_tipo, c.cif, c.direccion, c.cp, c.concello, c.telefono, c.email,
    (SELECT max(fecha) FROM trabajo_eventos e WHERE e.trabajo_id=t.id) ultimo_cambio,
    (SELECT json_build_object('tipo',d.tipo,'numero',d.numero,'base',d.base) FROM documentos d
       WHERE d.trabajo_id=t.id ORDER BY (d.tipo='factura') DESC, d.creado DESC LIMIT 1) doc
  FROM trabajos t JOIN clientes c ON c.id=t.cliente_id`;

router.get('/trabajos', wrap(async (req, res) => {
  const { tab = 'open', estado, tipo, q: texto, orden } = req.query;
  const w = []; const p = [];
  if (tab === 'open') w.push('t.cancelado_en IS NULL AND t.estado < 5');
  else if (tab === 'done') w.push('t.cancelado_en IS NULL AND t.estado = 5');
  else if (tab === 'cancel') w.push('t.cancelado_en IS NOT NULL');
  if (estado !== undefined && estado !== '') { p.push(+estado); w.push(`t.estado = $${p.length}`); }
  if (tipo && tipo !== 'Todos') { p.push(tipo); w.push(`t.tipo = $${p.length}`); }
  if (texto) { p.push('%' + texto + '%'); w.push(`(c.nombre ILIKE $${p.length} OR t.codigo ILIKE $${p.length} OR t.concello_obra ILIKE $${p.length} OR c.concello ILIKE $${p.length} OR t.descripcion ILIKE $${p.length})`); }
  const order = orden === 'az' ? 'cliente ASC' : 'coalesce(ultimo_cambio, creado) DESC';
  const r = await q(`SELECT * FROM (${TRABAJO_SELECT} ${w.length ? 'WHERE ' + w.join(' AND ') : ''}) x ORDER BY ${order} LIMIT 300`, p);
  res.json(r.rows);
}));

router.get('/trabajos/:id', wrap(async (req, res) => {
  const r = await q(`${TRABAJO_SELECT} WHERE t.id=$1`, [req.params.id]);
  if (!r.rows[0]) return res.status(404).json({ error: 'Non existe' });
  const ev = await q(`SELECT * FROM trabajo_eventos WHERE trabajo_id=$1 ORDER BY fecha DESC, id DESC`, [req.params.id]);
  const docs = await q(`SELECT * FROM documentos WHERE trabajo_id=$1 ORDER BY fecha DESC NULLS LAST, id DESC`, [req.params.id]);
  const arch = await q(`SELECT id,tipo,nombre,mime,tamano,creado,(miniatura IS NOT NULL) tiene_mini FROM archivos WHERE trabajo_id=$1 ORDER BY creado DESC`, [req.params.id]);
  res.json({ ...r.rows[0], eventos: ev.rows, documentos: docs.rows, archivos: arch.rows });
}));

async function nuevoCodigo() {
  const yy = String(new Date().getFullYear()).slice(2);
  const r = await q(`SELECT coalesce(max(split_part(codigo,'-',3)::int),0)+1 n FROM trabajos WHERE codigo LIKE $1`, [`T-${yy}-%`]);
  return `T-${yy}-${String(r.rows[0].n).padStart(3, '0')}`;
}

router.post('/trabajos', wrap(async (req, res) => {
  const b = req.body || {};
  let clienteId = b.cliente_id;
  if (!clienteId) {
    const c = b.cliente || {};
    if (!c.nombre) return res.status(400).json({ error: 'Falta el cliente' });
    const ins = await q(`INSERT INTO clientes (nombre,tipo,cif,direccion,cp,concello,telefono,email)
      VALUES ($1,$2,$3,$4,$5,$6,$7,$8)
      ON CONFLICT (cif,nombre) DO UPDATE SET actualizado=now() RETURNING id`,
      [c.nombre.trim(), c.tipo || 'Particular', (c.cif || '').trim().toUpperCase(), c.direccion || '', c.cp || '', c.concello || '', c.telefono || '', c.email || '']);
    clienteId = ins.rows[0].id;
  }
  const codigo = await nuevoCodigo();
  const t = await q(`INSERT INTO trabajos (codigo,cliente_id,tipo,descripcion,origen,direccion_obra,cp_obra,concello_obra,estado)
    VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9) RETURNING id`,
    [codigo, clienteId, b.tipo || 'Velux', b.descripcion || '', b.origen || '', b.direccion_obra || '', b.cp_obra || '', b.concello_obra || '', +b.estado || 0]);
  await q(`INSERT INTO trabajo_eventos (trabajo_id,tipo,estado,texto) VALUES ($1,'creado',$2,$3)`, [t.rows[0].id, +b.estado || 0, 'Trabajo creado']);
  res.json({ id: t.rows[0].id, codigo });
}));

const EDITABLES = ['tipo', 'descripcion', 'origen', 'direccion_obra', 'cp_obra', 'concello_obra', 'presupuesto_enviado', 'visita_en', 'montaje_en', 'notas'];
router.patch('/trabajos/:id', wrap(async (req, res) => {
  const b = req.body || {}; const sets = []; const p = [];
  EDITABLES.forEach((k) => { if (k in b) { p.push(b[k] === '' && /_en$|enviado/.test(k) ? null : b[k]); sets.push(`${k}=$${p.length}`); } });
  if (!sets.length) return res.json({ ok: true });
  p.push(req.params.id);
  await q(`UPDATE trabajos SET ${sets.join(',')}, actualizado=now() WHERE id=$${p.length}`, p);
  if ('presupuesto_enviado' in b && b.presupuesto_enviado) await q(`INSERT INTO trabajo_eventos (trabajo_id,tipo,texto) VALUES ($1,'nota','Presupuesto enviado')`, [req.params.id]);
  calendario.alCambiar(+req.params.id);
  res.json({ ok: true });
}));

router.post('/trabajos/:id/estado', wrap(async (req, res) => {
  const e = +(req.body || {}).estado;
  if (!(e >= 0 && e <= 5)) return res.status(400).json({ error: 'Estado non válido' });
  await q(`UPDATE trabajos SET estado=$1, cancelado_en=NULL, actualizado=now() WHERE id=$2`, [e, req.params.id]);
  await q(`INSERT INTO trabajo_eventos (trabajo_id,tipo,estado,texto) VALUES ($1,'estado',$2,$3)`, [req.params.id, e, 'Estado cambiado – ' + ESTADOS[e]]);
  if (e === 5) await q(`UPDATE documentos SET cobrada=true, cobrada_en=coalesce(cobrada_en,current_date) WHERE trabajo_id=$1 AND tipo='factura'`, [req.params.id]);
  calendario.alCambiar(+req.params.id);
  res.json({ ok: true });
}));

router.post('/trabajos/:id/cancelar', wrap(async (req, res) => {
  await q(`UPDATE trabajos SET cancelado_en=now(), actualizado=now() WHERE id=$1`, [req.params.id]);
  await q(`INSERT INTO trabajo_eventos (trabajo_id,tipo,texto) VALUES ($1,'cancelado',$2)`, [req.params.id, 'Trabajo cancelado' + ((req.body || {}).motivo ? ': ' + req.body.motivo : '')]);
  calendario.alCambiar(+req.params.id);
  res.json({ ok: true });
}));

// Vincular una factura/presupuesto ya existente (por número) o registrar uno nuevo
router.post('/trabajos/:id/documentos', wrap(async (req, res) => {
  const b = req.body || {};
  const tipo = b.tipo === 'presupuesto' ? 'presupuesto' : 'factura';
  const numero = b.numero ? normNumero(b.numero) : null;
  const t = await q(`SELECT cliente_id FROM trabajos WHERE id=$1`, [req.params.id]);
  if (!t.rows[0]) return res.status(404).json({ error: 'Non existe' });
  if (numero) {
    const ex = await q(`SELECT id FROM documentos WHERE tipo=$1 AND numero=$2`, [tipo, numero]);
    if (ex.rows[0]) {
      await q(`UPDATE documentos SET trabajo_id=$1 WHERE id=$2`, [req.params.id, ex.rows[0].id]);
      await q(`INSERT INTO trabajo_eventos (trabajo_id,tipo,texto) VALUES ($1,'documento',$2)`, [req.params.id, `${cap(tipo)} ${numero} vinculada`]);
      calendario.alCambiar(+req.params.id);
      return res.json({ ok: true, vinculado: true });
    }
  }
  await q(`INSERT INTO documentos (tipo,numero,fecha,base,total,origen,cliente_id,trabajo_id) VALUES ($1,$2,$3,$4,$5,$6,$7,$8)`,
    [tipo, numero, b.fecha || new Date().toISOString().slice(0, 10), num(b.base), num(b.total || b.base), b.origen || 'app', t.rows[0].cliente_id, req.params.id]);
  await q(`INSERT INTO trabajo_eventos (trabajo_id,tipo,texto) VALUES ($1,'documento',$2)`, [req.params.id, `${cap(tipo)}${numero ? ' ' + numero : ''} añadida`]);
  res.json({ ok: true });
}));
router.delete('/documentos/:id/vinculo', wrap(async (req, res) => {
  await q(`UPDATE documentos SET trabajo_id=NULL WHERE id=$1`, [req.params.id]);
  res.json({ ok: true });
}));

const cap = (s) => s.charAt(0).toUpperCase() + s.slice(1);
function normNumero(n) { return String(n).trim().replace(/\s+/g, '').replace('/', '-'); }

router.get('/siguiente-factura', wrap(async (_req, res) => {
  const yy = String(new Date().getFullYear()).slice(2);
  const r = await q(`SELECT coalesce(max(split_part(numero,'-',1)::int),0)+1 n FROM documentos WHERE tipo='factura' AND numero ~ $1`, [`^[0-9]+-${yy}$`]);
  res.json({ numero: `${r.rows[0].n}-${yy}` });
}));

// ── Clientes ────────────────────────────────────────────────────────────────
router.get('/clientes', wrap(async (req, res) => {
  const t = (req.query.q || '').trim();
  const r = await q(`SELECT c.*, (SELECT count(*)::int FROM documentos d WHERE d.cliente_id=c.id AND d.tipo='factura') nf,
      (SELECT coalesce(sum(base),0) FROM documentos d WHERE d.cliente_id=c.id AND d.tipo='factura') total
    FROM clientes c ${t ? `WHERE c.nombre ILIKE $1 OR c.cif ILIKE $1 OR c.direccion ILIKE $1 OR c.concello ILIKE $1 OR c.telefono ILIKE $1` : ''}
    ORDER BY c.nombre LIMIT ${t ? 40 : 400}`, t ? ['%' + t + '%'] : []);
  res.json(r.rows);
}));
router.get('/clientes/:id', wrap(async (req, res) => {
  const c = await q(`SELECT * FROM clientes WHERE id=$1`, [req.params.id]);
  if (!c.rows[0]) return res.status(404).json({ error: 'Non existe' });
  const d = await q(`SELECT * FROM documentos WHERE cliente_id=$1 ORDER BY fecha DESC NULLS LAST`, [req.params.id]);
  const t = await q(`SELECT id,codigo,tipo,estado,cancelado_en,descripcion,creado FROM trabajos WHERE cliente_id=$1 ORDER BY creado DESC`, [req.params.id]);
  res.json({ ...c.rows[0], documentos: d.rows, trabajos: t.rows });
}));
router.patch('/clientes/:id', wrap(async (req, res) => {
  const b = req.body || {}; const sets = []; const p = [];
  ['nombre', 'tipo', 'cif', 'direccion', 'cp', 'concello', 'telefono', 'email', 'notas'].forEach((k) => { if (k in b) { p.push(String(b[k])); sets.push(`${k}=$${p.length}`); } });
  if (!sets.length) return res.json({ ok: true });
  p.push(req.params.id);
  await q(`UPDATE clientes SET ${sets.join(',')}, actualizado=now() WHERE id=$${p.length}`, p);
  res.json({ ok: true });
}));

// ── Importar desde el Ficheiro de Clientes (public/clientes/index.html) ─────
// Se puede repetir: actualiza clientes y añade las facturas nuevas.
function fechaISO(s) { const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s || ''); return m ? `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}` : null; }
router.post('/sync-fichero', wrap(async (_req, res) => {
  const html = fs.readFileSync(path.join(__dirname, 'public', 'clientes', 'index.html'), 'utf8');
  const m = html.match(/const DATA\s*=\s*(\[[\s\S]*?\]);\s*\n/);
  if (!m) return res.status(500).json({ error: 'Non se atopou DATA no ficheiro' });
  const data = JSON.parse(m[1]);
  let nuevosC = 0, nuevasF = 0;
  const cli = await pool.connect();
  try {
    await cli.query('BEGIN');
    for (const c of data) {
      const r = await cli.query(`INSERT INTO clientes (nombre,tipo,cif,direccion,cp,concello,avisos,alias,origen)
        VALUES ($1,$2,$3,$4,$5,$6,$7,$8,'fichero')
        ON CONFLICT (cif,nombre) DO UPDATE SET tipo=EXCLUDED.tipo, direccion=EXCLUDED.direccion, cp=EXCLUDED.cp,
          concello=EXCLUDED.concello, avisos=EXCLUDED.avisos, alias=EXCLUDED.alias, actualizado=now()
        RETURNING id, (xmax = 0) nuevo`,
        [c.n, c.t, c.c || '', c.d || '', c.cp || '', c.m || '', JSON.stringify(c.av || []), JSON.stringify(c.alt || [])]);
      if (r.rows[0].nuevo) nuevosC++;
      for (const f of c.fs || []) {
        const fr = await cli.query(`INSERT INTO documentos (tipo,numero,fecha,base,total,origen,cliente_id,cobrada)
          VALUES ('factura',$1,$2,$3,$4,'word',$5,true) ON CONFLICT (tipo,numero) WHERE numero IS NOT NULL DO NOTHING RETURNING id`,
          [normNumero(f.n), fechaISO(f.f), f.b, f.t, r.rows[0].id]);
        if (fr.rows[0]) nuevasF++;
      }
    }
    await cli.query('COMMIT');
  } catch (e) { await cli.query('ROLLBACK'); throw e; } finally { cli.release(); }
  const tot = await q(`SELECT (SELECT count(*)::int FROM clientes) c, (SELECT count(*)::int FROM documentos WHERE tipo='factura') f, (SELECT coalesce(sum(base),0) FROM documentos WHERE tipo='factura') b`);
  res.json({ nuevosClientes: nuevosC, nuevasFacturas: nuevasF, clientes: tot.rows[0].c, facturas: tot.rows[0].f, base: num(tot.rows[0].b) });
}));


// ── Archivos: fotos y documentos ────────────────────────────────────────────
// Las fotos llegan ya reducidas desde el navegador (campo "file") con su miniatura (campo "mini").
router.post('/trabajos/:id/archivos', subida.fields([{ name: 'file', maxCount: 20 }, { name: 'mini', maxCount: 20 }]), wrap(async (req, res) => {
  const t = await q(`SELECT id, cliente_id, codigo FROM trabajos WHERE id=$1`, [req.params.id]);
  if (!t.rows[0]) return res.status(404).json({ error: 'Non existe' });
  const files = (req.files && req.files.file) || [];
  const minis = (req.files && req.files.mini) || [];
  const tipo = ['foto', 'presupuesto', 'factura', 'otro'].includes(req.body.tipo) ? req.body.tipo : 'foto';
  const creados = [];
  for (let i = 0; i < files.length; i++) {
    const f = files[i];
    const nombre = Buffer.from(f.originalname, 'latin1').toString('utf8');
    const a = await q(`INSERT INTO archivos (trabajo_id,cliente_id,tipo,nombre,mime,tamano,datos,miniatura) VALUES ($1,$2,$3,$4,$5,$6,$7,$8) RETURNING id`,
      [t.rows[0].id, t.rows[0].cliente_id, tipo, nombre, f.mimetype, f.size, f.buffer, minis[i] ? minis[i].buffer : null]);
    creados.push(a.rows[0].id);
    // Un Word de presupuesto o factura hecho fuera: se leen importes y se registra como documento
    if ((tipo === 'presupuesto' || tipo === 'factura') && /\.docx$/i.test(nombre)) {
      let d = {};
      try { d = await readPresuposto(f.buffer); } catch (e) { d = {}; }
      const numero = req.body.numero ? normNumero(req.body.numero) : (tipo === 'factura' ? (nombre.match(/(\d+)[-_ ](\d{2})/) || []).slice(1).join('-') || null : null);
      const base = num(String(d.base || '').replace(/[^0-9.]/g, '')), total = num(String(d.total || '').replace(/[^0-9.]/g, ''));
      await q(`INSERT INTO documentos (tipo,numero,fecha,base,total,origen,cliente_id,trabajo_id,lineas,iva_pct,archivo_id)
               VALUES ($1,$2,current_date,$3,$4,'word',$5,$6,$7,$8,$9)
               ON CONFLICT (tipo,numero) WHERE numero IS NOT NULL DO UPDATE SET trabajo_id=EXCLUDED.trabajo_id, archivo_id=EXCLUDED.archivo_id`,
        [tipo, numero, base, total || base, t.rows[0].cliente_id, t.rows[0].id, JSON.stringify(d.lineas || []), num(d.ivaPct || 21), a.rows[0].id]);
    }
  }
  const txt = tipo === 'foto' ? `${files.length} foto${files.length === 1 ? '' : 's'} añadida${files.length === 1 ? '' : 's'}` : `${cap(tipo)} subido: ${files.map((f) => f.originalname).join(', ')}`;
  await q(`INSERT INTO trabajo_eventos (trabajo_id,tipo,texto) VALUES ($1,'archivo',$2)`, [t.rows[0].id, txt]);
  res.json({ ok: true, ids: creados });
}));
router.get('/archivos/:id', wrap(async (req, res) => {
  const mini = req.query.mini === '1';
  const r = await q(`SELECT nombre,mime,${mini ? 'coalesce(miniatura,datos)' : 'datos'} datos FROM archivos WHERE id=$1`, [req.params.id]);
  if (!r.rows[0]) return res.status(404).end();
  const a = r.rows[0];
  res.setHeader('Content-Type', mini ? 'image/jpeg' : a.mime);
  res.setHeader('Cache-Control', 'private, max-age=31536000, immutable');
  if (req.query.descargar === '1') res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(a.nombre)}`);
  res.send(a.datos);
}));
router.delete('/archivos/:id', wrap(async (req, res) => {
  await q(`UPDATE documentos SET archivo_id=NULL WHERE archivo_id=$1`, [req.params.id]);
  await q(`DELETE FROM archivos WHERE id=$1`, [req.params.id]);
  res.json({ ok: true });
}));
// Lista para el script del Mac que copia todo a iCloud
router.get('/archivos-lista', wrap(async (req, res) => {
  const desde = req.query.desde || '1970-01-01';
  const r = await q(`SELECT a.id,a.tipo,a.nombre,a.mime,a.creado,t.codigo,c.nombre cliente,
      (SELECT d.numero FROM documentos d WHERE d.archivo_id=a.id LIMIT 1) numero,
      (SELECT d.origen FROM documentos d WHERE d.archivo_id=a.id LIMIT 1) origen
    FROM archivos a LEFT JOIN trabajos t ON t.id=a.trabajo_id LEFT JOIN clientes c ON c.id=a.cliente_id
    WHERE a.creado > $1 ORDER BY a.creado`, [desde]);
  res.json(r.rows);
}));
// Clave de larga duración para el script del Mac (se muestra una vez en Ajustes)
router.post('/clave-mac', (_req, res) => res.json({ token: sign({ u: 'mac', exp: Date.now() + 5 * 365 * 24 * 3600 * 1000 }) }));

// ── Presupuestos y facturas generados desde la ficha (mismo generador que el Xerador) ──
router.get('/trabajos/:id/siguientes', wrap(async (req, res) => {
  const yy = String(new Date().getFullYear()).slice(2);
  const f = await q(`SELECT coalesce(max(split_part(numero,'-',1)::int),0)+1 n FROM documentos WHERE tipo='factura' AND numero ~ $1`, [`^[0-9]+-${yy}$`]);
  const t = await q(`SELECT codigo FROM trabajos WHERE id=$1`, [req.params.id]);
  const np = await q(`SELECT count(*)::int n FROM documentos WHERE tipo='presupuesto' AND trabajo_id=$1`, [req.params.id]);
  const cod = t.rows[0] ? t.rows[0].codigo.replace(/^T-/, 'P-') : 'P';
  res.json({ factura: `${f.rows[0].n}-${yy}`, presupuesto: np.rows[0].n ? `${cod}-${np.rows[0].n + 1}` : cod });
}));

const TRABAJO_DOC_SQL = `SELECT t.*, c.nombre cnome, c.cif cnif, c.direccion cdir, c.cp ccp_cp, c.concello cconcello FROM trabajos t JOIN clientes c ON c.id=t.cliente_id WHERE t.id=$1`;
class ErrorHttp extends Error { constructor(st, m) { super(m); this.st = st; } }
// Monta el Word (mismo generador que el Xerador) a partir de los datos del formulario
async function armarDocumento(t, b, tipo, numero) {
  const lineas = (b.lineas || []).filter((l) => String(l.concepto || '').trim()).map((l) => ({
    ud: String(l.ud || '1'), concepto: String(l.concepto), precio: num(l.precio), subtotal: num(l.subtotal) || num(l.precio) * (parseFloat(l.ud) || 1) }));
  if (!lineas.length) throw new ErrorHttp(400, 'Añade al menos una línea');
  if (!numero) throw new ErrorHttp(400, 'Falta el número');
  const ivaPct = num(b.ivaPct === undefined ? 21 : b.ivaPct);
  const base = Math.round(lineas.reduce((a, l) => a + l.subtotal, 0) * 100) / 100;
  const ivaVal = Math.round(base * ivaPct) / 100;
  const total = Math.round((base + ivaVal) * 100) / 100;
  const payload = {
    num: tipo === 'factura' ? numero.replace('-', '/') : numero, fecha: b.fecha || new Date().toISOString().slice(0, 10),
    validez: b.validez || new Date(Date.now() + 90 * 864e5).toISOString().slice(0, 10),
    cnome: b.cnome || t.cnome, cnif: b.cnif || t.cnif,
    cdir: b.cdir || t.direccion_obra || t.cdir,
    ccp: b.ccp || [t.cp_obra || t.ccp_cp, t.concello_obra || t.cconcello].filter(Boolean).join(' '),
    objeto: b.objeto || t.descripcion, lineas, ivaPct: String(ivaPct),
    base: base.toFixed(2), ivaVal: ivaVal.toFixed(2), total: total.toFixed(2), notas: b.notas || '',
    modalidadePago: b.modalidadePago || 'unico', pagoUnico: b.pagoUnico || '', pago1: b.pago1 || '', pago2: b.pago2 || '', pago3: b.pago3 || '',
    banco: b.banco || 'bbva',
  };
  // Pagos por defecto: el material se cobra por adelantado (líneas que empiezan por «Material»), el resto al terminar
  if (tipo === 'presupuesto') {
    const k = 1 + ivaPct / 100, f2 = (v) => (Math.round(v * 100) / 100).toFixed(2);
    const mat = lineas.filter((l) => /^material/i.test(l.concepto)).reduce((a, l) => a + l.subtotal, 0) * k;
    if (payload.modalidadePago === 'unico' && !payload.pagoUnico) payload.pagoUnico = f2(total);
    if (payload.modalidadePago === 'dos' && !payload.pago1 && !payload.pago2) { const p1 = mat > 0 ? mat : total / 2; payload.pago1 = f2(p1); payload.pago2 = f2(total - p1); }
    if (payload.modalidadePago === 'tres' && !payload.pago1 && !payload.pago2 && !payload.pago3) { payload.pago1 = f2(mat > 0 ? mat : total / 3); const r = (total - payload.pago1) / 2; payload.pago2 = f2(r); payload.pago3 = f2(total - payload.pago1 - r); }
  }
  const buf = await Packer.toBuffer(tipo === 'factura' ? buildFactura(payload) : buildPresuposto(payload));
  const nombre = `${tipo === 'factura' ? 'Factura' : 'Presupuesto'}_${numero}.docx`;
  return { lineas, ivaPct, base, total, payload, buf, nombre };
}
const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const wrapDoc = (fn) => wrap(async (req, res) => { try { await fn(req, res); } catch (e) { if (e instanceof ErrorHttp) return res.status(e.st).json({ error: e.message }); throw e; } });

router.post('/trabajos/:id/generar', wrapDoc(async (req, res) => {
  const b = req.body || {};
  const tipo = b.tipo === 'factura' ? 'factura' : 'presupuesto';
  const t = (await q(TRABAJO_DOC_SQL, [req.params.id])).rows[0];
  if (!t) return res.status(404).json({ error: 'Non existe' });
  const numero = normNumero(b.num || '');
  if (numero) {
    const ex = await q(`SELECT id FROM documentos WHERE tipo=$1 AND numero=$2`, [tipo, numero]);
    if (ex.rows[0]) return res.status(409).json({ error: `Ya existe ${tipo === 'factura' ? 'la factura' : 'el presupuesto'} ${numero}` });
  }
  const { lineas, ivaPct, base, total, payload, buf, nombre } = await armarDocumento(t, b, tipo, numero);
  const a = await q(`INSERT INTO archivos (trabajo_id,cliente_id,tipo,nombre,mime,tamano,datos) VALUES ($1,$2,$3,$4,$5,$6,$7) RETURNING id`,
    [t.id, t.cliente_id, tipo, nombre, DOCX_MIME, buf.length, buf]);
  const d = await q(`INSERT INTO documentos (tipo,numero,fecha,base,total,origen,cliente_id,trabajo_id,lineas,iva_pct,datos,archivo_id)
     VALUES ($1,$2,$3,$4,$5,'app',$6,$7,$8,$9,$10,$11) RETURNING id`,
    [tipo, numero, payload.fecha, base, total, t.cliente_id, t.id, JSON.stringify(lineas), ivaPct, JSON.stringify(payload), a.rows[0].id]);
  await q(`INSERT INTO trabajo_eventos (trabajo_id,tipo,texto) VALUES ($1,'documento',$2)`, [t.id, `${cap(tipo)} ${numero} generad${tipo === 'factura' ? 'a' : 'o'} (${base.toFixed(2).replace('.', ',')} € base)`]);
  if (tipo === 'factura' && t.estado < 4 && !t.cancelado_en) {
    await q(`UPDATE trabajos SET estado=4, actualizado=now() WHERE id=$1`, [t.id]);
    await q(`INSERT INTO trabajo_eventos (trabajo_id,tipo,estado,texto) VALUES ($1,'estado',4,$2)`, [t.id, 'Estado cambiado – ' + ESTADOS[4]]);
  }
  if (tipo === 'presupuesto' && t.estado < 2 && !t.cancelado_en) {
    await q(`UPDATE trabajos SET estado=2, actualizado=now() WHERE id=$1`, [t.id]);
    await q(`INSERT INTO trabajo_eventos (trabajo_id,tipo,estado,texto) VALUES ($1,'estado',2,$2)`, [t.id, 'Estado cambiado – ' + ESTADOS[2]]);
  }
  calendario.alCambiar(t.id);
  res.json({ ok: true, documento_id: d.rows[0].id, archivo_id: a.rows[0].id, numero, nombre });
}));

// Modificar un presupuesto o factura ya hecho: se rehace el Word con los datos nuevos
router.put('/documentos/:id', wrapDoc(async (req, res) => {
  const b = req.body || {};
  const d0 = (await q(`SELECT * FROM documentos WHERE id=$1`, [req.params.id])).rows[0];
  if (!d0) return res.status(404).json({ error: 'Non existe' });
  if (!d0.trabajo_id) return res.status(400).json({ error: 'Este documento no está vinculado a un trabajo' });
  const tipo = d0.tipo;
  const t = (await q(TRABAJO_DOC_SQL, [d0.trabajo_id])).rows[0];
  const numero = normNumero(b.num || d0.numero || '');
  if (numero && numero !== d0.numero) {
    const ex = await q(`SELECT id FROM documentos WHERE tipo=$1 AND numero=$2 AND id<>$3`, [tipo, numero, d0.id]);
    if (ex.rows[0]) return res.status(409).json({ error: `Ya existe ${tipo === 'factura' ? 'la factura' : 'el presupuesto'} ${numero}` });
  }
  const { lineas, ivaPct, base, total, payload, buf, nombre } = await armarDocumento(t, b, tipo, numero);
  // archivo nuevo (así la copia del Mac se actualiza) y se borra el anterior
  const a = await q(`INSERT INTO archivos (trabajo_id,cliente_id,tipo,nombre,mime,tamano,datos) VALUES ($1,$2,$3,$4,$5,$6,$7) RETURNING id`,
    [t.id, t.cliente_id, tipo, nombre, DOCX_MIME, buf.length, buf]);
  await q(`UPDATE documentos SET numero=$1, fecha=$2, base=$3, total=$4, lineas=$5, iva_pct=$6, datos=$7, archivo_id=$8, origen='app' WHERE id=$9`,
    [numero, payload.fecha, base, total, JSON.stringify(lineas), ivaPct, JSON.stringify(payload), a.rows[0].id, d0.id]);
  if (d0.archivo_id) await q(`DELETE FROM archivos WHERE id=$1`, [d0.archivo_id]);
  await q(`INSERT INTO trabajo_eventos (trabajo_id,tipo,texto) VALUES ($1,'documento',$2)`, [t.id, `${cap(tipo)} ${numero} modificad${tipo === 'factura' ? 'a' : 'o'} (${base.toFixed(2).replace('.', ',')} € base)`]);
  calendario.alCambiar(t.id);
  res.json({ ok: true, documento_id: d0.id, archivo_id: a.rows[0].id, numero, nombre });
}));
router.get('/documentos/:id', wrap(async (req, res) => {
  const r = await q(`SELECT * FROM documentos WHERE id=$1`, [req.params.id]);
  if (!r.rows[0]) return res.status(404).json({ error: 'Non existe' });
  res.json(r.rows[0]);
}));
router.post('/documentos/:id/cobrada', wrap(async (req, res) => {
  const v = !!(req.body || {}).cobrada;
  await q(`UPDATE documentos SET cobrada=$1, cobrada_en=CASE WHEN $1 THEN current_date ELSE NULL END WHERE id=$2`, [v, req.params.id]);
  res.json({ ok: true });
}));

// ── Calendario de iCloud ────────────────────────────────────────────────────
router.get('/calendario', (_req, res) => res.json(calendario.estado()));
router.post('/calendario/sincronizar', wrap(async (_req, res) => res.json(await calendario.sincronizarTodo())));
router.get('/agenda', wrap(async (req, res) => res.json(await calendario.agenda(+req.query.dias || 21))));

// ── Exportar ────────────────────────────────────────────────────────────────
async function exportData() {
  const c = await q(`SELECT * FROM clientes ORDER BY nombre`);
  const d = await q(`SELECT * FROM documentos WHERE tipo='factura' ORDER BY fecha DESC NULLS LAST`);
  const dd = (x) => (x ? `${String(x.getDate()).padStart(2, '0')}/${String(x.getMonth() + 1).padStart(2, '0')}/${x.getFullYear()}` : '');
  return c.rows.map((x) => {
    const fs_ = d.rows.filter((f) => f.cliente_id === x.id);
    const b = fs_.reduce((a, f) => a + num(f.base), 0), t = fs_.reduce((a, f) => a + num(f.total), 0);
    return { n: x.nombre, t: x.tipo, c: x.cif, d: x.direccion, cp: x.cp, m: x.concello, tel: x.telefono, email: x.email,
      nf: fs_.length, b: Math.round(b * 100) / 100, tt: Math.round(t * 100) / 100,
      pf: fs_.length ? dd(fs_[fs_.length - 1].fecha) : '', uf: fs_.length ? dd(fs_[0].fecha) : '',
      ya: [...new Set(fs_.map((f) => f.fecha && String(f.fecha.getFullYear())).filter(Boolean))].sort(),
      alt: x.alias, av: x.avisos,
      fs: fs_.map((f) => ({ n: (f.numero || '').replace('-', '/'), f: dd(f.fecha), b: num(f.base), t: num(f.total) })) };
  });
}
router.get('/export/clientes.json', wrap(async (_req, res) => {
  res.setHeader('Content-Disposition', 'attachment; filename="clientes.json"');
  res.json(await exportData());
}));
router.get('/export/clientes.csv', wrap(async (_req, res) => {
  const rows = await exportData();
  const esc = (v) => '"' + String(v === undefined || v === null ? '' : v).replace(/"/g, '""') + '"';
  const cab = ['Cliente', 'Tipo', 'CIF/NIF', 'Dirección', 'C.P.', 'Concello', 'Teléfono', 'Email', 'Nº facturas', 'Total base', 'Total c/IVA', 'Primera', 'Última'];
  const lines = [cab.map(esc).join(';')].concat(rows.map((r) => [r.n, r.t, r.c, r.d, r.cp, r.m, r.tel, r.email, r.nf, String(r.b).replace('.', ','), String(r.tt).replace('.', ','), r.pf, r.uf].map(esc).join(';')));
  res.setHeader('Content-Type', 'text/csv; charset=utf-8');
  res.setHeader('Content-Disposition', 'attachment; filename="clientes.csv"');
  res.send('\ufeff' + lines.join('\r\n'));
}));

module.exports = router;
