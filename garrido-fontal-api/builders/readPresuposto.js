const mammoth = require('mammoth');

async function readPresuposto(buffer) {
  const result = await mammoth.extractRawText({ buffer });
  const lines  = result.value.split('\n').map(l => l.trim()).filter(Boolean);

  const data = {
    cnome: '', cnif: '', cdir: '', ccp: '',
    lineas: [], ivaPct: '21',
    base: '', ivaVal: '', total: '', notas: '',
  };

  const toNum = s => {
    const str = String(s).trim();
    // Spanish format "1.461,00" → "1461.00"
    if (str.includes(',')) return str.replace(/\./g, '').replace(',', '.');
    // Already decimal "1461.00" or "1767.81"
    return str.replace(/[^0-9.]/g, '');
  };

  // ── CLIENTE ───────────────────────────────────────────────────────────────
  // Formato real de mammoth (presupuesto en castellano):
  //   "Cliente:"
  //   "Carmen Pita Castro"
  //   "NIF/CIF:"
  //   "33808393A"
  //   "Dirección:"
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
    // Aceptar tanto "Dirección:" (nuevo) como "Enderezo:" (versiones antiguas)
    if (/^(Dirección|Direccion|Enderezo):$/i.test(l) && !data.cdir) {
      data.cdir = lines[i + 1] || '';
    }
    if (/^C\.P\.:$/i.test(l) && !data.ccp) {
      data.ccp = lines[i + 1] || '';
    }
  }

  // ── TOTALES ───────────────────────────────────────────────────────────────
  // Formato real (en castellano):
  //   "Base imponible:"
  //   "1.461,00 €"
  //   "IVA 21%:"
  //   "306,81 €"
  //   "TOTAL:"
  //   "1.767,81 €"

  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];

    // Aceptar "Base imponible:" (nuevo) y "Base impoñible" (antiguo bilingüe)
    if (/^Base\s+impo/i.test(l) && !data.base) {
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

    // Inline (todo en una línea)
    const bInline = l.match(/[Bb]ase[^0-9€]*([\d.]+,\d{2})\s*€/);
    if (bInline && !data.base) data.base = toNum(bInline[1]);

    const iInline = l.match(/IVA\s+(\d+)%[^0-9€]*([\d.]+,\d{2})\s*€/i);
    if (iInline && !data.ivaVal) { data.ivaPct = iInline[1]; data.ivaVal = toNum(iInline[2]); }

    const tInline = l.match(/^TOTAL[^0-9€]*([\d.]+,\d{2})\s*€/i);
    if (tInline && !data.total) data.total = toNum(tInline[1]);
  }

  // ── PARTIDAS ──────────────────────────────────────────────────────────────
  // Formato real (presupuesto en castellano):
  //   "1"                                         <- ud sola
  //   "Ventana Velux GGL SK06..."                 <- concepto único
  //   "1.461,00 €"                                <- precio
  //   "1.461,00 €"                                <- subtotal

  let inItems  = false;
  let state    = null; // null | 'gotUd' | 'gotConcepto'
  let current  = null;

  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];

    // Cabeceras de tabla — entrar en modo items
    if (/^(Ud\.|Concepto|Precio|Prezo|Total)/i.test(l)) { inItems = true; continue; }
    // Salir al ver secciones siguientes
    if (/^Notas\s*[:\/]/i.test(l) || /^Base\s+impo/i.test(l) || /^Forma\s+de\s+pago/i.test(l) || /^Objeto/i.test(l)) {
      inItems = false; current = null; state = null;
    }
    if (!inItems) continue;

    // Precio "1.461,00 €" o "1461.00 €"
    const priceM = l.match(/^([\d.,]+)\s*€\s*$/);

    // Ud sola (solo dígitos)
    if (/^\d+$/.test(l) && !priceM) {
      current = { ud: l, concepto: '', precio: '0', subtotal: '0' };
      state = 'gotUd';
      continue;
    }

    // Precio
    if (priceM && current) {
      const val = toNum(priceM[1]);
      if (current.precio === '0') {
        current.precio   = val;
        current.subtotal = val;
      } else if (current.subtotal === current.precio) {
        current.subtotal = val;
        if (current.concepto) data.lineas.push(current);
        current = null; state = null;
      }
      continue;
    }

    // Em-dash = precio incluido
    if (/^[—–]$/.test(l)) continue;

    // Concepto principal (1ª línea de texto tras la Ud)
    if (current && state === 'gotUd' && l.length > 2) {
      current.concepto = l;
      state = 'gotConcepto';
      continue;
    }

    // Línea adicional del concepto (descripción larga, multilínea)
    // Ya no traducimos: solo añadimos como continuación
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
