// ── Conversión de Word (.docx) a PDF ─────────────────────────────────────────
// Usa LibreOffice instalado en el servidor (variable RAILPACK_DEPLOY_APT_PACKAGES en Railway).
// Opcional: si existe GOTENBERG_URL se usa ese servicio en su lugar.
const { execFile } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const BINARIOS = ['soffice', 'libreoffice'];
let binario; // undefined = sin comprobar · null = no hay · string = ruta/nombre

function buscarBinario() {
  if (binario !== undefined) return Promise.resolve(binario);
  return new Promise((ok) => {
    const probar = (i) => {
      if (i >= BINARIOS.length) { binario = null; return ok(null); }
      execFile('which', [BINARIOS[i]], (e, out) => (!e && out.trim() ? (binario = out.trim(), ok(binario)) : probar(i + 1)));
    };
    probar(0);
  });
}

async function estado() {
  const gotenberg = !!process.env.GOTENBERG_URL;
  const local = !!(await buscarBinario());
  return { disponible: gotenberg || local, motor: gotenberg ? 'Gotenberg' : local ? 'LibreOffice' : null };
}

// Una conversión cada vez: LibreOffice no admite dos instancias con el mismo perfil
let cola = Promise.resolve();
function enCola(fn) { const r = cola.then(fn, fn); cola = r.catch(() => {}); return r; }

async function conGotenberg(buffer) {
  const fd = new FormData();
  fd.append('files', new Blob([buffer], { type: DOCX_MIME }), 'documento.docx');
  const r = await fetch(process.env.GOTENBERG_URL.replace(/\/$/, '') + '/forms/libreoffice/convert', { method: 'POST', body: fd, signal: AbortSignal.timeout(90000) });
  if (!r.ok) throw new Error('Conversor PDF: error ' + r.status);
  return Buffer.from(await r.arrayBuffer());
}

async function conLibreOffice(buffer) {
  const bin = await buscarBinario();
  if (!bin) throw new Error('El servidor no tiene LibreOffice instalado (falta RAILPACK_DEPLOY_APT_PACKAGES)');
  const dir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'gfpdf-'));
  try {
    await fs.promises.writeFile(path.join(dir, 'documento.docx'), buffer);
    await new Promise((ok, ko) => {
      execFile(bin, ['--headless', '--norestore', '--nolockcheck', `-env:UserInstallation=file://${path.join(dir, 'perfil')}`, '--convert-to', 'pdf', '--outdir', dir, path.join(dir, 'documento.docx')],
        { timeout: 90000, env: { ...process.env, HOME: dir } }, (e, _o, err) => (e ? ko(new Error('LibreOffice: ' + (err || e.message).toString().slice(0, 200))) : ok()));
    });
    return await fs.promises.readFile(path.join(dir, 'documento.pdf'));
  } finally { fs.promises.rm(dir, { recursive: true, force: true }).catch(() => {}); }
}

function docxAPdf(buffer) {
  return enCola(() => (process.env.GOTENBERG_URL ? conGotenberg(buffer) : conLibreOffice(buffer)));
}

module.exports = { docxAPdf, estado };
