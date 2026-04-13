const mammoth = require('mammoth');

async function readPresuposto(buffer) {
  const result = await mammoth.extractRawText({ buffer });
  const lines  = result.value.split('\n').map(l => l.trim()).filter(Boolean);

  const data = {
    cnome: '', cnif: '', cdir: '', ccp: '',
    lineas: [], ivaPct: '21',
    base: '', ivaVal: '', total: '', notas: '',
  };

  const toNum = s => String(s).replace(/\./g, '').replace(',', '.');

  // ── CLIENTE ───────────────────────────────────────────────────────────────
  // Formato real de mammoth con esta app:
  //   "Cliente:"       <- línea sola
  //   "Carmen Pita Castro"
  //   "NIF/CIF:"
  //   "33808393A"
  //   "Enderezo:"
  //   "Rúa Armónica Nº 42, 2ºB"
  //   "C.P.:"
  //   "27002 Lugo"

  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];

    if (/^Cliente:$/i.test(l) && !data.cnome) {
      data.cnome = lines[i + 1] || '';
    }
    if (/^NIF\/CIF:$/i.test(l) && !data.cnif) {
      data.cnif = lines[i + 1] || '';
    }
    if (/^Enderezo:$/i.test(l) && !data.cdir) {
      data.cdir = lines[i + 1] || '';
    }
    if (/^C\.P\.:$/i.test(l) && !data.ccp) {
      data.ccp = lines[i + 1] || '';
    }
  }

  // ── TOTAIS ────────────────────────────────────────────────────────────────
  // Formato real:
  //   "Base impoñible / Base imponible:"
  //   "1461.00"          <- línea siguiente (sin €)
  //   "IVA 21%:"
  //   "306.81"
  //   "TOTAL:"
  //   "1767.81"

  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];

    if (/^Base impo/i.test(l) && !data.base) {
      const next = lines[i + 1] || '';
      const m = next.match(/^[\d.,]+/);
      if (m) data.base = toNum(m[0]);
    }

    const ivaM = l.match(/^IVA\s+(\d+)%/i);
    if (ivaM && !data.ivaVal) {
      data.ivaPct = ivaM[1];
      const next = lines[i + 1] || '';
      const m = next.match(/^[\d.,]+/);
      if (m) data.ivaVal = toNum(m[0]);
    }

    if (/^TOTAL:$/i.test(l) && !data.total) {
      const next = lines[i + 1] || '';
      const m = next.match(/^[\d.,]+/);
      if (m) data.total = toNum(m[0]);
    }

    // También por si vienen en la misma línea con €
    const bInline = l.match(/[Bb]ase[^0-9€]*([\d.]+,\d{2})\s*€/);
    if (bInline && !data.base) data.base = toNum(bInline[1]);

    const iInline = l.match(/IVA\s+(\d+)%[^0-9€]*([\d.]+,\d{2})\s*€/i);
    if (iInline && !data.ivaVal) { data.ivaPct = iInline[1]; data.ivaVal = toNum(iInline[2]); }

    const tInline = l.match(/^TOTAL[^0-9€]*([\d.]+,\d{2})\s*€/i);
    if (tInline && !data.total) data.total = toNum(tInline[1]);
  }

  // ── PARTIDAS ──────────────────────────────────────────────────────────────
  // Formato real de mammoth:
  //   "1"                                         <- ud sola
  //   "Montaxe de VELUX SK06..."                  <- concepto galego
  //   "Montaje de VELUX SK06..."                  <- concepto castellano (itálica)
  //   "1461.00 €"                                 <- precio
  //   "1461.00 €"                                 <- subtotal

  let inItems  = false;
  let state    = null; // null | 'gotUd' | 'gotConcepto'
  let current  = null;

  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];

    if (/^(Ud\.|Concepto|Prezo|Precio|Total)/i.test(l)) { inItems = true; continue; }
    if (/^Notas\s*\//i.test(l) || /^Base\s+impo/i.test(l) || /^Forma\s+de\s+pago/i.test(l)) {
      inItems = false; current = null; state = null;
    }
    if (!inItems) continue;

    // Precio suelto "1461.00 €" o "1.461,00 €"
    const priceM = l.match(/^([\d.,]+)\s*€\s*$/);

    // Ud sola (solo dígitos)
    if (/^\d+$/.test(l) && !priceM) {
      if (current && state === 'gotUd') {
        // Ud anterior sin concepto, saltar
      }
      current = { ud: l, concepto: '', precio: '0', subtotal: '0' };
      state = 'gotUd';
      continue;
    }

    // Precio: "1461.00 €"
    if (priceM && current) {
      const val = toNum(priceM[1]);
      if (current.precio === '0') {
        current.precio   = val;
        current.subtotal = val;
      } else if (current.subtotal === current.precio) {
        current.subtotal = val;
        // Item completo — añadir si tiene concepto
        if (current.concepto) data.lineas.push(current);
        current = null; state = null;
      }
      continue;
    }

    // Em-dash = precio incluido
    if (/^[—–]$/.test(l)) continue;

    // Concepto
    if (current && state === 'gotUd' && l.length > 2) {
      current.concepto = l;
      state = 'gotConcepto';
      continue;
    }

    // Segunda línea del concepto (traducción castellano)
    if (current && state === 'gotConcepto' && l.length > 2 && !priceM) {
      if (!current.concepto.includes('\n')) {
        current.concepto += '\n' + l;
      }
      continue;
    }
  }

  // Añadir último item si quedó pendiente
  if (current && current.concepto && current.precio !== '0') {
    data.lineas.push(current);
  }

  return data;
}

module.exports = { readPresuposto };
