const mammoth = require('mammoth');

async function readPresuposto(buffer) {
  const result = await mammoth.extractRawText({ buffer });
  const lines  = result.value.split('\n').map(l => l.trim()).filter(Boolean);

  const data = {
    cnome: '', cnif: '', cdir: '', ccp: '',
    lineas: [], ivaPct: '21',
    base: '', ivaVal: '', total: '', notas: '',
  };

  const toNum = s => s.replace(/\./g, '').replace(',', '.');

  // ── CLIENTE ───────────────────────────────────────────────────────────────
  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];

    if (/^Cliente:/i.test(l)) {
      const m = l.replace(/^Cliente:\s*/i, '').replace(/\s{2,}.*$/, '').trim();
      if (m) data.cnome = m;
      for (let j = i + 1; j < Math.min(i + 5, lines.length); j++) {
        const nx = lines[j];
        if (/^Enderezo:|^Válido|^Obxecto|^NIF|^CIF/i.test(nx)) break;
        if (/^\d{5}\s/.test(nx)) { data.ccp = nx; break; }
        if (/^Rúa|^Avda|^Praza|^Calle|^C\//i.test(nx)) data.cdir = nx;
      }
    }

    if (/^Enderezo:/i.test(l)) {
      const addr = l.replace(/^Enderezo:\s*/i, '').replace(/\s{2,}.*$/, '').trim();
      if (/^\d{5}/.test(addr)) data.ccp  = addr;
      else if (addr)            data.cdir = addr;
    }

    if (!data.cnif && /^[A-Z]-?\d{6,8}$|^\d{8}[A-Z]$|^[A-Z]\d{7}[A-Z]$/.test(l)) data.cnif = l;
    if (!data.ccp  && /^\d{5}\s+\w/.test(l)) data.ccp = l;
  }

  // ── TOTAIS ────────────────────────────────────────────────────────────────
  for (const l of lines) {
    const bM = l.match(/[Bb]ase[^0-9€]*([\d.]+,\d{2})\s*€/);
    if (bM && !data.base) data.base = toNum(bM[1]);

    const iM = l.match(/IVA\s+(\d+)%[^0-9€]*([\d.]+,\d{2})\s*€/i);
    if (iM) { data.ivaPct = iM[1]; data.ivaVal = toNum(iM[2]); }

    const tM = l.match(/^TOTAL[^0-9€]*([\d.]+,\d{2})\s*€/i);
    if (tM && !data.total) data.total = toNum(tM[1]);
  }

  // ── PARTIDAS ──────────────────────────────────────────────────────────────
  let inItems = false;
  let current = null;

  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];

    if (/Concepto\s*\/\s*Descripci/i.test(l))                              { inItems = true; continue; }
    if (/^Notas\s*\//i.test(l) || /^Base\s+impo/i.test(l) || /^Forma\s+de\s+pago/i.test(l)) { inItems = false; current = null; }
    if (!inItems) continue;
    if (/^Prezo\s+Ud|^Precio\s+Ud|^Total\s+s\/IVA|^Ud\s*\./i.test(l))     continue;
    if (/^Piso\s+\d|^Planta\s+\d/i.test(l))                               { current = null; continue; }

    // Tab-separated row
    const tp = l.split('\t').map(p => p.trim()).filter(Boolean);
    if (tp.length >= 2 && /^\d+$/.test(tp[0])) {
      const pM = (tp[2]||'').match(/([\d.]+,\d{2})/);
      const sM = (tp[3]||'').match(/([\d.]+,\d{2})/);
      current = { ud: tp[0], concepto: tp[1],
        precio:   pM ? toNum(pM[1]) : '0',
        subtotal: sM ? toNum(sM[1]) : (pM ? toNum(pM[1]) : '0') };
      data.lineas.push(current);
      continue;
    }

    // Space-separated: "1  Concepto  1.461,00 €  1.461,00 €"
    const m2 = l.match(/^(\d+)\s{2,}(.+?)\s{2,}([\d.]+,\d{2})\s*€\s+([\d.]+,\d{2})\s*€/);
    if (m2) { current = { ud:m2[1], concepto:m2[2].trim(), precio:toNum(m2[3]), subtotal:toNum(m2[4]) }; data.lineas.push(current); continue; }

    const m3 = l.match(/^(\d+)\s{2,}(.+?)\s{2,}([\d.]+,\d{2})\s*€/);
    if (m3) { current = { ud:m3[1], concepto:m3[2].trim(), precio:toNum(m3[3]), subtotal:toNum(m3[3]) }; data.lineas.push(current); continue; }

    const m4 = l.match(/^(\d+)\s{2,}(.+)$/);
    if (m4 && !/^\d{1,2}\/\d{2}\/\d{4}/.test(m4[2])) {
      current = { ud:m4[1], concepto:m4[2].trim(), precio:'0', subtotal:'0' };
      data.lineas.push(current);
      continue;
    }

    // Continuation
    if (current) {
      const pO = l.match(/^([\d.]+,\d{2})\s*€\s*$/);
      if (pO) { if (current.precio==='0') { current.precio=toNum(pO[1]); current.subtotal=current.precio; } continue; }
      if (/^[—–]$/.test(l) || /^[-─═]+$/.test(l)) continue;
      if (!current.concepto.includes('\n') && l.length > 3 && !/^[A-ZÁÉÍÓÚ].*:$/.test(l)) {
        current.concepto += '\n' + l;
      }
    }
  }

  // IVA fallback
  if (data.base && data.total && !data.ivaVal) {
    const b = parseFloat(data.base), t = parseFloat(data.total);
    if (b && t) { data.ivaVal = (t-b).toFixed(2); data.ivaPct = String(Math.round(((t-b)/b)*100)); }
  }

  return data;
}

module.exports = { readPresuposto };
