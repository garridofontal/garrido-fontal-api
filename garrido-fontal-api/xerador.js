// ── Xerador: PIN, fichero de clientes, archivo de documentos y backup ──────────
// Se monta desde server.js:
//   const xerador = require('./xerador');
//   app.use(xerador.gate);            // ANTES de express.static
//   app.use('/api', xerador.router);  // /api/clientes, /api/xerador/*
//   migrate().then(() => xerador.init()) ...
const express = require('express');
const crypto  = require('crypto');
const fs      = require('fs');
const path    = require('path');
const multer  = require('multer');
const { q, pool } = require('./db');

const PIN = process.env.APP_PIN || '';
const COOKIE = 'gf_sesion';
const DIAS_SESION = 90;
const subida = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });

// ── Sesión por cookie (HMAC del PIN) ──────────────────────────────────────────
function firma() {
  return crypto.createHmac('sha256', PIN).update('garrido-fontal-sesion').digest('hex');
}
function cookies(req) {
  const out = {};
  (req.headers.cookie || '').split(';').forEach(p => {
    const i = p.indexOf('='); if (i < 0) return;
    out[p.slice(0, i).trim()] = decodeURIComponent(p.slice(i + 1).trim());
  });
  return out;
}
function iguales(a, b) {
  const A = Buffer.from(String(a)), B = Buffer.from(String(b));
  return A.length === B.length && crypto.timingSafeEqual(A, B);
}
function autorizado(req) {
  if (!PIN) return true;                                   // sin APP_PIN: sin protección (aviso en init)
  const pinCab = req.headers['x-pin'];
  if (pinCab && iguales(pinCab, PIN)) return true;         // scripts del Mac
  const c = cookies(req)[COOKIE];
  return !!c && iguales(c, firma());
}

// Rutas protegidas por PIN (estáticas y API)
const PROTEGIDAS = [/^\/$/, /^\/index\.html$/, /^\/clientes\/?$/, /^\/clientes\/index\.html$/, /^\/api\/clientes/, /^\/api\/xerador\/(?!login)/];
const LIBRES = [/^\/clientes\/(manifest\.webmanifest|sw\.js|favicon\.png|icon-.*\.png|apple-touch-icon\.png)$/];

function gate(req, res, next) {
  const p = req.path;
  if (LIBRES.some(r => r.test(p))) return next();
  if (!PROTEGIDAS.some(r => r.test(p))) return next();
  if (autorizado(req)) return next();
  if (p.startsWith('/api/')) return res.status(401).json({ error: 'PIN requerido' });
  res.status(401).type('html').send(paginaLogin(p));
}

function paginaLogin(destino) {
  return `<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Garrido Fontal — Acceso</title>
<style>body{margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;background:#F7F9F9;font-family:system-ui,sans-serif;color:#1A2A2A}
.c{background:#fff;border:1px solid #D8DDDD;border-radius:10px;padding:28px;width:300px;text-align:center;box-shadow:0 8px 24px rgba(0,0,0,.06)}
h1{font-size:18px;margin:0 0 4px}p{font-size:13px;color:#6B7B7B;margin:0 0 18px}
input{width:100%;box-sizing:border-box;padding:12px;font-size:22px;letter-spacing:6px;text-align:center;border:1px solid #D8DDDD;border-radius:8px}
button{width:100%;margin-top:12px;padding:12px;background:#2E4A4A;color:#fff;border:0;border-radius:8px;font-size:15px;font-weight:600;cursor:pointer}
.e{color:#993C1D;font-size:13px;margin-top:10px;display:none}</style></head><body>
<form class="c" onsubmit="return entrar(event)"><h1>Garrido Fontal SLU</h1><p>Introduce el PIN para continuar</p>
<input id="pin" type="password" inputmode="numeric" autocomplete="current-password" autofocus required>
<button>Entrar</button><div class="e" id="e">PIN incorrecto</div></form>
<script>async function entrar(ev){ev.preventDefault();const r=await fetch('/api/xerador/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pin:document.getElementById('pin').value})});
if(r.ok){location.href=${JSON.stringify(destino || '/')};}else{document.getElementById('e').style.display='block';document.getElementById('pin').value='';}return false;}</script>
</body></html>`;
}

// ── Router ────────────────────────────────────────────────────────────────────
const router = express.Router();
const wrap = fn => (req, res) => fn(req, res).catch(e => { console.error(e); res.status(500).json({ error: e.message }); });

router.post('/xerador/login', (req, res) => {
  const pin = String((req.body || {}).pin || '');
  if (!PIN || !iguales(pin, PIN)) return res.status(401).json({ error: 'PIN incorrecto' });
  res.setHeader('Set-Cookie', `${COOKIE}=${firma()}; Path=/; Max-Age=${DIAS_SESION * 86400}; HttpOnly; Secure; SameSite=Lax`);
  res.json({ ok: true });
});
router.post('/xerador/logout', (_req, res) => {
  res.setHeader('Set-Cookie', `${COOKIE}=; Path=/; Max-Age=0; HttpOnly; Secure; SameSite=Lax`);
  res.json({ ok: true });
});

// ── Fichero de clientes (misma tabla que el organizador) ──────────────────────
const TIPO_POR_CIF = l => /^H/i.test(l) ? 'Comunidade' : /^[PQRS]/i.test(l) ? 'Organismo' : /^[ABCDEFGJNUVW]/i.test(l) ? 'Empresa' : 'Particular';

function leerFicheroDATA() {
  try {
    const html = fs.readFileSync(path.join(__dirname, 'public', 'clientes', 'index.html'), 'utf8');
    const m = html.match(/const DATA\s*=\s*(\[[\s\S]*?\]);\s*\n/);
    return m ? JSON.parse(m[1]) : [];
  } catch (e) { console.error('Fichero de clientes:', e.message); return []; }
}

async function syncFicheroSiVacio() {
  if (!pool) return;
  const r = await q('SELECT count(*)::int AS n FROM clientes');
  if (r.rows[0].n > 0) return;
  const data = leerFicheroDATA();
  for (const c of data) {
    await q("INSERT INTO clientes (nombre,tipo,cif,direccion,cp,concello,avisos,alias,origen) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,'fichero') ON CONFLICT (cif,nombre) DO NOTHING",
      [c.n, c.t || 'Particular', c.c || '', c.d || '', c.cp || '', c.m || '', JSON.stringify(c.av || []), JSON.stringify(c.alt || [])]);
  }
  console.log('Fichero de clientes cargado en Postgres: ' + data.length);
}

router.get('/clientes', wrap(async (req, res) => {
  const t = String(req.query.q || '').trim();
  const todos = req.query.todos === '1';
  if (pool) {
    const r = await q('SELECT id,nombre,tipo,cif,direccion,cp,concello,telefono,email,origen,creado,actualizado FROM clientes ' +
      (t ? 'WHERE nombre ILIKE $1 OR cif ILIKE $1 OR direccion ILIKE $1 OR concello ILIKE $1 ' : '') +
      'ORDER BY nombre' + (todos ? '' : ' LIMIT ' + (t ? 30 : 400)), t ? ['%' + t + '%'] : []);
    return res.json(r.rows);
  }
  const tl = t.toLowerCase();
  res.json(leerFicheroDATA()
    .filter(c => !tl || [c.n, c.c, c.d, c.m].some(v => String(v || '').toLowerCase().includes(tl)))
    .slice(0, todos ? undefined : 30)
    .map(c => ({ id: null, nombre: c.n, tipo: c.t, cif: c.c || '', direccion: c.d || '', cp: c.cp || '', concello: c.m || '' })));
}));

router.post('/clientes', wrap(async (req, res) => {
  const b = req.body || {};
  const nombre = String(b.nombre || '').trim();
  if (!nombre) return res.status(400).json({ error: 'Falta el nombre' });
  if (!pool) return res.status(503).json({ error: 'Sin base de datos' });
  const cif = String(b.cif || '').trim().toUpperCase().replace(/[\s-]/g, '');
  const tipo = cif ? TIPO_POR_CIF(cif) : 'Particular';
  const r = await q("INSERT INTO clientes (nombre,tipo,cif,direccion,cp,concello,origen) VALUES ($1,$2,$3,$4,$5,$6,'xerador') " +
    "ON CONFLICT (cif,nombre) DO UPDATE SET direccion=COALESCE(NULLIF(EXCLUDED.direccion,''),clientes.direccion), cp=COALESCE(NULLIF(EXCLUDED.cp,''),clientes.cp), " +
    "concello=COALESCE(NULLIF(EXCLUDED.concello,''),clientes.concello), actualizado=now() RETURNING id, nombre, (xmax = 0) AS nuevo",
    [nombre, tipo, cif, String(b.direccion || '').trim(), String(b.cp || '').trim(), String(b.concello || '').trim()]);
  res.json(r.rows[0]);
}));

// ── Archivo de documentos generados por el Xerador ────────────────────────────
async function crearTablas() {
  if (!pool) return;
  await q(`CREATE TABLE IF NOT EXISTS xerador_docs (
    id             SERIAL PRIMARY KEY,
    tipo           TEXT NOT NULL,
    numero         TEXT NOT NULL DEFAULT '',
    fecha          DATE,
    cliente_id     INTEGER,
    cliente_nombre TEXT NOT NULL DEFAULT '',
    cif            TEXT NOT NULL DEFAULT '',
    base           NUMERIC(12,2),
    iva_pct        NUMERIC(5,2),
    total          NUMERIC(12,2),
    datos          JSONB NOT NULL DEFAULT '{}',
    nombre_archivo TEXT NOT NULL,
    mime           TEXT NOT NULL,
    contenido      BYTEA NOT NULL,
    creado         TIMESTAMPTZ NOT NULL DEFAULT now()
  );
  CREATE INDEX IF NOT EXISTS xerador_docs_creado ON xerador_docs (creado);`);
}

router.post('/xerador/guardar-doc', subida.single('file'), wrap(async (req, res) => {
  if (!pool) return res.status(503).json({ error: 'Sin base de datos' });
  if (!req.file) return res.status(400).json({ error: 'Falta el archivo' });
  const b = req.body || {};
  const datos = (() => { try { return JSON.parse(b.datos || '{}'); } catch { return {}; } })();
  const num = v => (v === '' || v === undefined || v === null || isNaN(+v)) ? null : +v;
  const r = await q(`INSERT INTO xerador_docs (tipo,numero,fecha,cliente_id,cliente_nombre,cif,base,iva_pct,total,datos,nombre_archivo,mime,contenido)
    VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13) RETURNING id, creado`,
    [String(b.tipo || 'documento'), String(b.numero || ''), b.fecha || null, num(b.cliente_id), String(b.cliente_nombre || ''), String(b.cif || ''),
     num(b.base), num(b.iva_pct), num(b.total), JSON.stringify(datos), req.file.originalname, req.file.mimetype, req.file.buffer]);
  res.json(r.rows[0]);
}));

// Listado (sin binario) — ?desde=ISO para incrementales
router.get('/xerador/docs', wrap(async (req, res) => {
  if (!pool) return res.json([]);
  const desde = req.query.desde ? new Date(req.query.desde) : null;
  const r = await q(`SELECT id,tipo,numero,fecha,cliente_id,cliente_nombre,cif,base,iva_pct,total,nombre_archivo,mime,octet_length(contenido) AS bytes,creado
    FROM xerador_docs ${desde && !isNaN(desde) ? 'WHERE creado > $1' : ''} ORDER BY creado`, desde && !isNaN(desde) ? [desde] : []);
  res.json(r.rows);
}));

router.get('/xerador/docs/:id/archivo', wrap(async (req, res) => {
  const r = await q('SELECT nombre_archivo, mime, contenido FROM xerador_docs WHERE id=$1', [req.params.id]);
  if (!r.rows[0]) return res.status(404).json({ error: 'No existe' });
  res.setHeader('Content-Disposition', `attachment; filename="${r.rows[0].nombre_archivo.replace(/"/g, '')}"`);
  res.setHeader('Content-Type', r.rows[0].mime);
  res.send(r.rows[0].contenido);
}));

// ── Backup completo en JSON (todas las tablas, sin columnas binarias) ─────────
router.get('/xerador/backup', wrap(async (_req, res) => {
  if (!pool) return res.status(503).json({ error: 'Sin base de datos' });
  const tablas = (await q(`SELECT table_name FROM information_schema.tables WHERE table_schema='public' AND table_type='BASE TABLE' ORDER BY table_name`)).rows.map(r => r.table_name);
  const out = { generado: new Date().toISOString(), tablas: {} };
  for (const t of tablas) {
    const cols = (await q(`SELECT column_name, data_type FROM information_schema.columns WHERE table_schema='public' AND table_name=$1 ORDER BY ordinal_position`, [t])).rows;
    const sel = cols.map(c => c.data_type === 'bytea' ? `octet_length("${c.column_name}") AS "${c.column_name}__bytes"` : `"${c.column_name}"`).join(',');
    const filas = (await q(`SELECT ${sel} FROM "${t}"`)).rows;
    out.tablas[t] = { columnas: cols.map(c => c.column_name), filas };
  }
  res.setHeader('Content-Disposition', `attachment; filename="backup-garrido-fontal-${out.generado.slice(0, 10)}.json"`);
  res.json(out);
}));

async function init() {
  if (!PIN) console.warn('AVISO: APP_PIN no está definido — las apps quedan sin protección por PIN');
  try { await crearTablas(); } catch (e) { console.error('xerador_docs:', e.message); }
  try { await syncFicheroSiVacio(); } catch (e) { console.error('syncFicheroSiVacio:', e.message); }
}

module.exports = { gate, router, init };
