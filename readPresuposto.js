const mammoth = require('mammoth');

/**
 * Extracts structured data from a Garrido Fontal presuposto .docx
 * Returns an object compatible with the factura form fields
 */
async function readPresuposto(buffer) {
  const result = await mammoth.extractRawText({ buffer });
  const text = result.value;
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);

  const data = {
    cnome: '', cnif: '', cdir: '', ccp: '',
    lineas: [],
    ivaPct: '21',
    base: '', ivaVal: '', total: '',
    notas: '',
    rawLines: lines,
  };

  // ── Client data ────────────────────────────────────────────────────────────
  // The meta table renders as text blocks; we look for patterns
  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];

    // NIF/CIF pattern: letter(s) + dash + digits, or 8 digits + letter
    if (!data.cnif && /^[A-Z]-?\d{7,8}$|^\d{8}[A-Z]$|^[A-Z]\d{7}[A-Z]$/.test(l)) {
      data.cnif = l;
    }

    // Postal code + city (5 digits at start)
    if (!data.ccp && /^\d{5}\s+\w/.test(l)) {
      data.ccp = l;
    }

    // Street address
    if (!data.cdir && /^R[úu]a\s|^Avda?\.|^Praza\s|^Calle\s|^C\/|^Av\./i.test(l)) {
      data.cdir = l;
    }
  }

  // Client name — usually the line before NIF or a long proper-noun line
  // Look for "Cliente:" label then following lines
  const clienteIdx = lines.findIndex(l => /^cliente/i.test(l));
  if (clienteIdx >= 0) {
    // Scan next lines for the name (skip short labels)
    for (let i = clienteIdx + 1; i < Math.min(clienteIdx + 6, lines.length); i++) {
      const l = lines[i];
      if (l.length > 3 && !/^NIF|^CIF|^ORZAMENTO|^PRESUPUESTO|^\d{5}/i.test(l) && !data.cnif?.startsWith(l)) {
        data.cnome = l;
        break;
      }
    }
  }

  // ── Totals ─────────────────────────────────────────────────────────────────
  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];

    // Base impoñible: "4.383,00€" or "4.383,00 €"
    const baseMatch = l.match(/^[Bb]ase[^\d]*([\d.,]+)\s*€/);
    if (baseMatch) data.base = baseMatch[1].replace('.', '').replace(',', '.');

    // IVA line: "IVA 21% 920,43€"
    const ivaMatch = l.match(/^IVA\s+(\d+)%[^\d]*([\d.,]+)\s*€/i);
    if (ivaMatch) {
      data.ivaPct = ivaMatch[1];
      data.ivaVal = ivaMatch[2].replace('.', '').replace(',', '.');
    }

    // TOTAL: "TOTAL 5.303,43 €"
    const totalMatch = l.match(/^TOTAL[^\d]*([\d.,]+)\s*€/i);
    if (totalMatch) data.total = totalMatch[1].replace('.', '').replace(',', '.');
  }

  // ── Line items ─────────────────────────────────────────────────────────────
  // Items appear as: number | concept text | price | subtotal
  // We look for lines with a currency amount at the end
  const priceRe = /([\d.,]+)\s*€\s*$/;
  let inItems = false;

  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];

    // Start of items section
    if (/^Concepto\s*\/\s*Descripci/i.test(l)) { inItems = true; continue; }
    if (/^Notas\s*\/\s*Notas/i.test(l)) { inItems = false; continue; }
    if (/^Base\s+impo/i.test(l)) { inItems = false; continue; }

    if (!inItems) continue;

    // Line that starts with a number (ud) and has a price
    const udMatch = l.match(/^(\d+)\s+(.+?)\s+([\d.,]+)\s*€\s+([\d.,]+)\s*€\s*$/);
    if (udMatch) {
      data.lineas.push({
        ud: udMatch[1],
        concepto: udMatch[2].trim(),
        precio: udMatch[3].replace('.', '').replace(',', '.'),
        subtotal: udMatch[4].replace('.', '').replace(',', '.'),
      });
      continue;
    }

    // Concept-only line (continuation)
    if (data.lineas.length > 0 && !priceRe.test(l) && l.length > 5
        && !/^—$|^Ud\.$|^IVA|^Base|^TOTAL|^Forma|^\d+%/.test(l)) {
      // Append to last concept as second line (will appear as italic in factura)
      const last = data.lineas[data.lineas.length - 1];
      if (!last.concepto.includes('\n')) {
        last.concepto += '\n' + l;
      }
    }
  }

  // ── IVA from rate if not found ─────────────────────────────────────────────
  if (data.base && data.total && !data.ivaVal) {
    const b = parseFloat(data.base);
    const t = parseFloat(data.total);
    if (b && t) {
      data.ivaVal = (t - b).toFixed(2);
      const rate = Math.round((t - b) / b * 100);
      data.ivaPct = String(rate);
    }
  }

  return data;
}

module.exports = { readPresuposto };
