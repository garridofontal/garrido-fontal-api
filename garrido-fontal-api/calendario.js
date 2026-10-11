// ── Calendario de iCloud (CalDAV) en los dos sentidos ─────────────────────────
// App → calendario: visitas, montajes, recordatorios de cobro y seguimiento de presupuestos.
// Calendario → app: si se mueve o borra un evento en el iPhone/Mac, la ficha se actualiza.
// Variables en Railway: ICLOUD_USER, ICLOUD_APP_PASSWORD (contraseña específica de app).
const { q, pool } = require('./db');

const NOMBRE_CAL = process.env.ICLOUD_CALENDARIO || 'Trabajos Garrido Fontal';
const APP_URL = 'https://garrido-fontal-api-production.up.railway.app/trabajos/';
const TZ = 'Europe/Madrid';
const DIAS_COBRO = 15, DIAS_SEGUIMIENTO = 7;
const DURACION = { visita: 60, montaje: 300 }; // minutos

const estado = { configurado: false, conectado: false, calendario: null, error: null, ultimoSync: null, ultimoPull: null };
let client = null, calendar = null, conectando = null;

async function migrar() {
  await q(`CREATE TABLE IF NOT EXISTS eventos_cal (
    id          SERIAL PRIMARY KEY,
    trabajo_id  INTEGER NOT NULL REFERENCES trabajos(id) ON DELETE CASCADE,
    tipo        TEXT NOT NULL,
    uid         TEXT UNIQUE NOT NULL,
    url         TEXT,
    etag        TEXT,
    inicio      TIMESTAMPTZ,
    todo_dia    BOOLEAN NOT NULL DEFAULT false,
    descartado  BOOLEAN NOT NULL DEFAULT false,
    firma       TEXT,
    actualizado TIMESTAMPTZ NOT NULL DEFAULT now(),
    UNIQUE (trabajo_id, tipo)
  )`);
}

async function conectar() {
  if (calendar) return calendar;
  estado.configurado = !!(process.env.ICLOUD_USER && process.env.ICLOUD_APP_PASSWORD);
  if (!estado.configurado) { estado.error = 'Faltan ICLOUD_USER e ICLOUD_APP_PASSWORD en Railway'; return null; }
  if (conectando) return conectando;
  conectando = (async () => {
    try {
      const { createDAVClient } = require('tsdav');
      client = await createDAVClient({
        serverUrl: 'https://caldav.icloud.com',
        credentials: { username: process.env.ICLOUD_USER.trim(), password: process.env.ICLOUD_APP_PASSWORD.trim() },
        authMethod: 'Basic', defaultAccountType: 'caldav',
      });
      const buscar = async () => (await client.fetchCalendars()).find((c) => String(c.displayName || '').trim().toLowerCase() === NOMBRE_CAL.toLowerCase());
      let cal = await buscar();
      if (!cal) {
        const todos = await client.fetchCalendars();
        const home = (client.account && client.account.homeUrl) || (todos[0] && new URL('..', todos[0].url.endsWith('/') ? todos[0].url : todos[0].url + '/').href);
        const url = new URL('trabajos-garrido-fontal/', home).href;
        try { await client.makeCalendar({ url, props: { displayname: NOMBRE_CAL, 'I:calendar-color': '#2E4A4AFF' } }); } catch (e) { /* se reintenta abajo */ }
        cal = await buscar();
      }
      if (!cal) throw new Error(`No existe el calendario «${NOMBRE_CAL}». Créalo en la app Calendario (cuenta iCloud) con ese nombre exacto.`);
      calendar = cal; estado.conectado = true; estado.calendario = cal.displayName; estado.error = null;
      return cal;
    } catch (e) {
      client = null; calendar = null; estado.conectado = false;
      estado.error = /401|Unauthorized|invalid credentials/i.test(e.message) ? 'iCloud rechaza el usuario o la contraseña específica de app' : e.message;
      return null;
    } finally { conectando = null; }
  })();
  return conectando;
}

// ── Fechas ──────────────────────────────────────────────────────────────────
const p2 = (n) => String(n).padStart(2, '0');
const icsUTC = (d) => `${d.getUTCFullYear()}${p2(d.getUTCMonth() + 1)}${p2(d.getUTCDate())}T${p2(d.getUTCHours())}${p2(d.getUTCMinutes())}00Z`;
const icsDate = (d) => `${d.getUTCFullYear()}${p2(d.getUTCMonth() + 1)}${p2(d.getUTCDate())}`;
function fechaMadrid(d) { // YYYY-MM-DD en hora de Madrid
  const f = new Intl.DateTimeFormat('en-CA', { timeZone: TZ, year: 'numeric', month: '2-digit', day: '2-digit' }).format(d);
  return f;
}
function diaUTC(isoDia) { const [y, m, d] = isoDia.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)); }
function sumarDias(d, n) { return new Date(d.getTime() + n * 864e5); }
function offsetMin(tz, fecha) {
  const parts = Object.fromEntries(new Intl.DateTimeFormat('en-US', { timeZone: tz, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' })
    .formatToParts(fecha).filter((p) => p.type !== 'literal').map((p) => [p.type, +p.value]));
  return (Date.UTC(parts.year, parts.month - 1, parts.day, parts.hour, parts.minute, parts.second) - fecha.getTime()) / 60000;
}
function localAUtc(y, mo, d, h, mi, s, tz) {
  const guess = new Date(Date.UTC(y, mo - 1, d, h, mi, s));
  let off = 0;
  try { off = offsetMin(tz, guess); off = offsetMin(tz, new Date(guess.getTime() - off * 60000)); } catch (e) { off = offsetMin(TZ, guess); }
  return new Date(guess.getTime() - off * 60000);
}
// Lee DTSTART de un VEVENT → { fecha: Date, todoDia }
function leerInicio(ics) {
  const txt = ics.replace(/\r?\n[ \t]/g, '');
  const ev = txt.split('BEGIN:VEVENT')[1] || '';
  const m = ev.match(/\nDTSTART([^:\n]*):([0-9TZ]+)/);
  if (!m) return null;
  const params = m[1], v = m[2];
  if (/VALUE=DATE/.test(params) || v.length === 8) return { fecha: diaUTC(`${v.slice(0, 4)}-${v.slice(4, 6)}-${v.slice(6, 8)}`), todoDia: true };
  const y = +v.slice(0, 4), mo = +v.slice(4, 6), d = +v.slice(6, 8), h = +v.slice(9, 11), mi = +v.slice(11, 13), s = +v.slice(13, 15) || 0;
  if (v.endsWith('Z')) return { fecha: new Date(Date.UTC(y, mo - 1, d, h, mi, s)), todoDia: false };
  const tz = (params.match(/TZID=([^;:]+)/) || [])[1] || TZ;
  return { fecha: localAUtc(y, mo, d, h, mi, s, tz.replace(/"/g, '')), todoDia: false };
}

// ── Qué eventos debe tener cada trabajo ─────────────────────────────────────
async function eventosDeseados(t) {
  const out = [];
  if (t.cancelado_en) return out;
  const dir = [t.direccion_obra || t.direccion, [t.cp_obra || t.cp, t.concello_obra || t.concello].filter(Boolean).join(' ')].filter(Boolean).join(', ');
  const base = { dir, desc: [`${t.codigo} · ${t.tipo}`, t.descripcion, t.telefono ? 'Tel: ' + t.telefono : '', dir, APP_URL + '#/t/' + t.id].filter(Boolean).join('\n') };
  if (t.visita_en && t.estado <= 1) out.push({ ...base, tipo: 'visita', titulo: `Visita · ${t.cliente}`, inicio: new Date(t.visita_en), todoDia: false, campo: 'visita_en' });
  if (t.montaje_en && t.estado <= 3) out.push({ ...base, tipo: 'montaje', titulo: `Montaje · ${t.cliente}`, inicio: new Date(t.montaje_en), todoDia: false, campo: 'montaje_en' });
  const hoy = diaUTC(fechaMadrid(new Date()));
  if (t.estado === 2 && t.presupuesto_enviado) {
    let d = sumarDias(diaUTC(fechaMadrid(new Date(t.presupuesto_enviado))), DIAS_SEGUIMIENTO); if (d < hoy) d = sumarDias(hoy, 1);
    out.push({ ...base, tipo: 'seguimiento', titulo: `Seguimiento presupuesto · ${t.cliente}`, inicio: d, todoDia: true, recordatorio: true });
  }
  if (t.estado === 4) {
    const f = await q(`SELECT numero, fecha, total FROM documentos WHERE trabajo_id=$1 AND tipo='factura' ORDER BY fecha DESC NULLS LAST LIMIT 1`, [t.id]);
    if (f.rows[0]) {
      let d = sumarDias(diaUTC(fechaMadrid(f.rows[0].fecha || new Date())), DIAS_COBRO); if (d < hoy) d = sumarDias(hoy, 1);
      const tot = parseFloat(f.rows[0].total || 0).toFixed(2).replace('.', ',');
      out.push({ ...base, tipo: 'cobro', titulo: `Cobrar factura ${f.rows[0].numero || ''} · ${t.cliente} (${tot} €)`, inicio: d, todoDia: true, recordatorio: true });
    }
  }
  return out;
}
const esc = (s) => String(s || '').replace(/\\/g, '\\\\').replace(/\n/g, '\\n').replace(/,/g, '\\,').replace(/;/g, '\\;');
function ics(uid, e) {
  const fin = e.todoDia ? sumarDias(e.inicio, 1) : new Date(e.inicio.getTime() + (DURACION[e.tipo] || 60) * 60000);
  return ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Garrido Fontal//Trabajos//ES', 'CALSCALE:GREGORIAN', 'BEGIN:VEVENT',
    `UID:${uid}`, `DTSTAMP:${icsUTC(new Date())}`,
    e.todoDia ? `DTSTART;VALUE=DATE:${icsDate(e.inicio)}` : `DTSTART:${icsUTC(e.inicio)}`,
    e.todoDia ? `DTEND;VALUE=DATE:${icsDate(fin)}` : `DTEND:${icsUTC(fin)}`,
    `SUMMARY:${esc(e.titulo)}`, e.dir ? `LOCATION:${esc(e.dir)}` : null, `DESCRIPTION:${esc(e.desc)}`, `URL:${APP_URL}`,
    e.todoDia ? null : 'BEGIN:VALARM\r\nACTION:DISPLAY\r\nDESCRIPTION:Recordatorio\r\nTRIGGER:-PT60M\r\nEND:VALARM',
    'END:VEVENT', 'END:VCALENDAR'].filter(Boolean).join('\r\n') + '\r\n';
}

const TRABAJO_SQL = `SELECT t.*, c.nombre cliente, c.telefono, c.direccion, c.cp, c.concello FROM trabajos t JOIN clientes c ON c.id=t.cliente_id`;

// ── App → calendario ────────────────────────────────────────────────────────
async function pushTrabajo(id) {
  const cal = await conectar(); if (!cal) return;
  const r = await q(`${TRABAJO_SQL} WHERE t.id=$1`, [id]); const t = r.rows[0];
  const actuales = (await q(`SELECT * FROM eventos_cal WHERE trabajo_id=$1`, [id])).rows;
  const deseados = t ? await eventosDeseados(t) : [];
  for (const e of deseados) {
    const ex = actuales.find((a) => a.tipo === e.tipo);
    if (ex && ex.descartado) continue;                       // lo borró en el calendario: no se vuelve a crear
    if (ex && e.recordatorio && ex.inicio) { e.inicio = new Date(ex.inicio); }  // respeta si lo movió en el calendario
    const uid = ex ? ex.uid : `gf-${id}-${e.tipo}@garridofontal`;
    const firma = `${e.titulo}|${e.inicio.toISOString()}|${e.todoDia}|${e.dir}|${e.desc}`;
    if (ex && ex.firma === firma && ex.url) continue;
    const data = ics(uid, e);
    let url = ex && ex.url, etag = ex && ex.etag;
    if (url) {
      let res = await client.updateCalendarObject({ calendarObject: { url, data, etag: etag || undefined } });
      if (res.status === 412) res = await client.updateCalendarObject({ calendarObject: { url, data } });
      if (res.status === 404) url = null; else if (!res.ok) throw new Error('iCloud update ' + res.status);
      else etag = res.headers.get('etag');
    }
    if (!url) {
      const filename = uid.replace(/[^a-z0-9-]/gi, '_') + '.ics';
      const res = await client.createCalendarObject({ calendar: cal, filename, iCalString: data });
      if (!res.ok && res.status !== 412) throw new Error('iCloud create ' + res.status);
      url = new URL(filename, cal.url.endsWith('/') ? cal.url : cal.url + '/').href; etag = res.headers.get('etag');
    }
    await q(`INSERT INTO eventos_cal (trabajo_id,tipo,uid,url,etag,inicio,todo_dia,firma,actualizado) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,now())
      ON CONFLICT (trabajo_id,tipo) DO UPDATE SET url=$4, etag=$5, inicio=$6, todo_dia=$7, firma=$8, actualizado=now()`,
      [id, e.tipo, uid, url, etag, e.inicio, e.todoDia, firma]);
  }
  for (const a of actuales) {                                // eventos que ya no tocan (estado avanzado, fecha borrada, cancelado)
    if (deseados.find((e) => e.tipo === a.tipo)) continue;
    if (a.url && !a.descartado) { try { await client.deleteCalendarObject({ calendarObject: { url: a.url, etag: a.etag || undefined } }); } catch (e) { /* ya no estaba */ } }
    await q(`DELETE FROM eventos_cal WHERE id=$1`, [a.id]);
  }
  estado.ultimoSync = new Date();
}
function alCambiar(id) { pushTrabajo(id).catch((e) => { estado.error = e.message; console.error('Calendario:', e.message); }); }

// ── Calendario → app ────────────────────────────────────────────────────────
async function pull() {
  const cal = await conectar(); if (!cal) return { cambios: 0 };
  const objs = await client.fetchCalendarObjects({ calendar: cal });
  const porUrl = new Map(objs.map((o) => [o.url, o]));
  const porUid = new Map(objs.map((o) => [((o.data || '').match(/UID:([^\r\n]+)/) || [])[1], o]));
  const filas = (await q(`SELECT ec.*, t.codigo FROM eventos_cal ec JOIN trabajos t ON t.id=ec.trabajo_id WHERE NOT ec.descartado`)).rows;
  let cambios = 0;
  for (const f of filas) {
    const o = porUrl.get(f.url) || porUid.get(f.uid);
    if (!o) {                                               // borrado en el calendario
      if (f.tipo === 'visita' || f.tipo === 'montaje') {
        await q(`UPDATE trabajos SET ${f.tipo}_en=NULL, actualizado=now() WHERE id=$1`, [f.trabajo_id]);
        await q(`DELETE FROM eventos_cal WHERE id=$1`, [f.id]);
      } else await q(`UPDATE eventos_cal SET descartado=true, actualizado=now() WHERE id=$1`, [f.id]);
      await q(`INSERT INTO trabajo_eventos (trabajo_id,tipo,texto) VALUES ($1,'calendario',$2)`, [f.trabajo_id, `${nombreTipo(f.tipo)} ${f.tipo === 'visita' ? 'borrada' : 'borrado'} en el calendario`]);
      cambios++; continue;
    }
    if (o.etag && o.etag === f.etag) continue;
    const ini = leerInicio(o.data || ''); if (!ini) continue;
    const antes = f.inicio ? new Date(f.inicio).getTime() : 0;
    if (Math.abs(ini.fecha.getTime() - antes) < 60000) { await q(`UPDATE eventos_cal SET etag=$1 WHERE id=$2`, [o.etag, f.id]); continue; }
    if (f.tipo === 'visita' || f.tipo === 'montaje') await q(`UPDATE trabajos SET ${f.tipo}_en=$1, actualizado=now() WHERE id=$2`, [ini.fecha, f.trabajo_id]);
    await q(`UPDATE eventos_cal SET inicio=$1, todo_dia=$2, etag=$3, firma=NULL, url=$4, actualizado=now() WHERE id=$5`, [ini.fecha, ini.todoDia, o.etag, o.url, f.id]);
    const txt = ini.todoDia ? ini.fecha.toISOString().slice(0, 10).split('-').reverse().join('/')
      : new Intl.DateTimeFormat('es-ES', { timeZone: TZ, day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }).format(ini.fecha);
    await q(`INSERT INTO trabajo_eventos (trabajo_id,tipo,texto) VALUES ($1,'calendario',$2)`, [f.trabajo_id, `${nombreTipo(f.tipo)} ${f.tipo === 'visita' ? 'movida' : 'movido'} en el calendario al ${txt}`]);
    cambios++;
    alCambiar(f.trabajo_id);                                 // reescribe el evento con nuestro formato y etag nuevo
  }
  estado.ultimoPull = new Date();
  return { cambios };
}
const nombreTipo = (t) => ({ visita: 'Visita', montaje: 'Montaje', cobro: 'Recordatorio de cobro', seguimiento: 'Seguimiento de presupuesto' }[t] || t);

async function sincronizarTodo() {
  const cal = await conectar(); if (!cal) return { ok: false, error: estado.error };
  const p = await pull();
  const ids = (await q(`SELECT id FROM trabajos WHERE cancelado_en IS NULL AND estado < 5 UNION SELECT trabajo_id FROM eventos_cal`)).rows.map((r) => r.id);
  for (const id of ids) await pushTrabajo(id);
  return { ok: true, trabajos: ids.length, cambiosDesdeCalendario: p.cambios };
}

// Agenda de los próximos días (para la pantalla de inicio), con o sin iCloud
async function agenda(dias = 21) {
  const r = await q(`${TRABAJO_SQL} WHERE t.cancelado_en IS NULL AND t.estado < 5`);
  const hasta = Date.now() + dias * 864e5, desde = Date.now() - 864e5;
  const guard = (await q(`SELECT trabajo_id, tipo, inicio, descartado FROM eventos_cal`)).rows;
  const out = [];
  for (const t of r.rows) {
    for (const e of await eventosDeseados(t)) {
      const g = guard.find((x) => x.trabajo_id === t.id && x.tipo === e.tipo);
      if (g && g.descartado) continue;
      if (g && e.recordatorio && g.inicio) e.inicio = new Date(g.inicio);
      const ms = e.inicio.getTime();
      if (ms >= desde && ms <= hasta) out.push({ trabajo_id: t.id, codigo: t.codigo, cliente: t.cliente, tipo: e.tipo, titulo: e.titulo, inicio: e.inicio, todo_dia: e.todoDia });
    }
  }
  return out.sort((a, b) => a.inicio - b.inicio);
}

function arrancar() {
  if (!pool) return;
  migrar().then(() => {
    if (!(process.env.ICLOUD_USER && process.env.ICLOUD_APP_PASSWORD)) { estado.error = 'Faltan ICLOUD_USER e ICLOUD_APP_PASSWORD en Railway'; return; }
    setTimeout(() => sincronizarTodo().catch((e) => { estado.error = e.message; }), 5000);
    setInterval(() => pull().catch((e) => { estado.error = e.message; console.error('Calendario pull:', e.message); }), 5 * 60 * 1000);
  }).catch((e) => console.error('Calendario: migración', e));
}

function _mock(c, cal) { client = c; calendar = cal; estado.conectado = true; estado.configurado = true; }
module.exports = { _mock, migrar, pushTrabajo, arrancar, alCambiar, pull, sincronizarTodo, agenda, estado: () => ({ ...estado, nombre: NOMBRE_CAL }), leerInicio };
