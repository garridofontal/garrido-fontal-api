const {
  Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, ImageRun, Footer
} = require('docx');
const path = require('path');
const fs   = require('fs');

// ── Date formatter ────────────────────────────────────────────────────────────
function formatDate(dateStr) {
  if (!dateStr) return '—';
  const parts = dateStr.split('-');
  if (parts.length === 3) return `${parts[2]}/${parts[1]}/${parts[0]}`;
  return dateStr;
}

// ── Number formatter ──────────────────────────────────────────────────────────
function formatNum(val) {
  if (val === undefined || val === null || val === '—' || val === '') return '—';
  const n = parseFloat(String(val).replace(',', '.'));
  if (isNaN(n)) return String(val);
  return n.toLocaleString('es-ES', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// ── Colors ────────────────────────────────────────────────────────────────────
const C = {
  DARK:     '2E4A4A',
  LIGHT_BG: 'F4F7F7',
  WHITE:    'FFFFFF',
  GRAY_TEXT:'555555',
  GRAY_LINE:'CCCCCC',
};

const LOGO = fs.readFileSync(path.join(__dirname, '..', 'logo.png'));

// ── Border helpers ────────────────────────────────────────────────────────────
function noBorder() {
  const n = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  return { top: n, bottom: n, left: n, right: n };
}
function thinBorder(color = C.GRAY_LINE) {
  const b = { style: BorderStyle.SINGLE, size: 4, color };
  return { top: b, bottom: b, left: b, right: b };
}

// ── Paragraph helpers ─────────────────────────────────────────────────────────
function para(text, opts = {}) {
  return new Paragraph({
    alignment: opts.align ?? AlignmentType.LEFT,
    spacing:   opts.spacing ?? { before: 0, after: 0 },
    children: [new TextRun({
      text, font: 'Arial',
      size:    opts.size  ?? 20,
      bold:    opts.bold  ?? false,
      color:   opts.color ?? '000000',
      italics: opts.italic ?? false,
    })],
  });
}
function emptyPara(pts = 80) {
  return new Paragraph({ spacing: { before: 0, after: pts }, children: [] });
}

// ── Table cell helpers ────────────────────────────────────────────────────────
function labelCell(text, w) {
  return new TableCell({
    borders: noBorder(),
    width: { size: w, type: WidthType.DXA },
    shading: { fill: C.LIGHT_BG, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 150, right: 150 },
    children: [para(text, { size: 18, bold: true, color: C.DARK })],
  });
}
function valueCell(text, w, opts = {}) {
  return new TableCell({
    borders: noBorder(),
    width: { size: w, type: WidthType.DXA },
    margins: { top: 80, bottom: 80, left: 150, right: 150 },
    children: [para(text, { size: 18, color: '222222', ...opts })],
  });
}

// ── Header (logo izquierda + datos empresa derecha) ───────────────────────────
// Ancho total página = 9326 twips (A4 – márgenes)
// Col logo: 2200 | Col datos: 7126
function buildHeader() {
  const LOGO_COL  = 2200;
  const DATA_COL  = 7126;
  const TOTAL     = LOGO_COL + DATA_COL; // 9326

  return new Table({
    width: { size: TOTAL, type: WidthType.DXA },
    columnWidths: [LOGO_COL, DATA_COL],
    borders: {
      top:     { style: BorderStyle.NONE },
      bottom:  { style: BorderStyle.NONE },
      left:    { style: BorderStyle.NONE },
      right:   { style: BorderStyle.NONE },
      insideH: { style: BorderStyle.NONE },
      insideV: { style: BorderStyle.NONE },
    },
    rows: [new TableRow({
      children: [
        // ─── Celda logo ───────────────────────────────────────────────────
        new TableCell({
          borders:         noBorder(),
          width:           { size: LOGO_COL, type: WidthType.DXA },
          verticalAlign:   VerticalAlign.CENTER,
          margins:         { top: 0, bottom: 0, left: 0, right: 200 },
          children: [new Paragraph({
            alignment: AlignmentType.LEFT,
            spacing:   { before: 0, after: 0 },
            children:  [new ImageRun({
              data:           LOGO,
              transformation: { width: 120, height: 120 },
              type:           'png',
            })],
          })],
        }),
        // ─── Celda datos empresa ──────────────────────────────────────────
        new TableCell({
          borders:       noBorder(),
          width:         { size: DATA_COL, type: WidthType.DXA },
          verticalAlign: VerticalAlign.BOTTOM,
          margins:       { top: 0, bottom: 10, left: 0, right: 0 },
          children: [
            para('Garrido Fontal SLU', { size: 28, bold: true, color: C.DARK }),
            para('Rúa Bispo Doutor Balanza, 14-3º · 27002 Lugo',        { size: 18, color: C.GRAY_TEXT }),
            para('garridofontalslu@gmail.com · WhatsApp: 687 398 413',  { size: 18, color: C.GRAY_TEXT }),
            para('CIF: B27203520',                                        { size: 18, color: C.GRAY_TEXT }),
          ],
        }),
      ],
    })],
  });
}

// ── Separator ─────────────────────────────────────────────────────────────────
function buildSeparator() {
  return new Paragraph({
    spacing: { before: 120, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.DARK, space: 1 } },
    children: [],
  });
}

// ── Footer ────────────────────────────────────────────────────────────────────
function buildFooter() {
  return new Footer({
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing:   { before: 0, after: 0 },
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.DARK, space: 4 } },
      children: [new TextRun({
        text: 'Garrido Fontal SLU · Rúa Bispo Doutor Balanza 14-3º, 27002 Lugo · garridofontalslu@gmail.com · 687 398 413',
        font: 'Arial', size: 16, color: C.GRAY_TEXT,
      })],
    })],
  });
}

// ── Totals table ──────────────────────────────────────────────────────────────
function buildTotals(base, ivaPct, ivaVal, total) {
  function row(label, value, isBold = false) {
    const displayVal = (value === '—' || value === undefined || value === null)
      ? '—' : formatNum(value) + ' €';
    return new TableRow({ children: [
      new TableCell({
        borders: noBorder(),
        width: { size: 7826, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 150, right: 150 },
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing:   { before: 0, after: 0 },
          children:  [new TextRun({ text: label, font: 'Arial', size: isBold ? 22 : 20, bold: isBold, color: isBold ? C.DARK : '444444' })],
        })],
      }),
      new TableCell({
        borders: thinBorder(isBold ? C.DARK : C.GRAY_LINE),
        width:   { size: 1500, type: WidthType.DXA },
        shading: { fill: isBold ? C.DARK : C.WHITE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 150, right: 150 },
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing:   { before: 0, after: 0 },
          children:  [new TextRun({ text: displayVal, font: 'Arial', size: isBold ? 22 : 20, bold: isBold, color: isBold ? C.WHITE : '222222' })],
        })],
      }),
    ]});
  }
  return new Table({
    width: { size: 9326, type: WidthType.DXA },
    columnWidths: [7826, 1500],
    rows: [
      row('Base impoñible / Base imponible:', base),
      row(`IVA ${ivaPct}%:`, ivaVal),
      row('TOTAL:', total, true),
    ],
  });
}

// ── Items table cells ─────────────────────────────────────────────────────────
function headerItemCell(text, w, align = AlignmentType.CENTER) {
  return new TableCell({
    width:   { size: w, type: WidthType.DXA },
    shading: { fill: C.DARK, type: ShadingType.CLEAR },
    borders: thinBorder(C.DARK),
    margins: { top: 100, bottom: 100, left: 150, right: 150 },
    children: [new Paragraph({
      alignment: align,
      spacing:   { before: 0, after: 0 },
      children:  [new TextRun({ text, font: 'Arial', size: 18, bold: true, color: C.WHITE })],
    })],
  });
}

function conceptCell(lines, w, isAlt = false) {
  const fill = isAlt ? 'EEF3F3' : C.WHITE;
  const children = [];
  lines.forEach((l, i) => {
    const parts = String(l.text || '').split('\n');
    parts.forEach((part, pi) => {
      children.push(new Paragraph({
        spacing: { before: (i === 0 && pi === 0) ? 0 : 40, after: 0 },
        children: [new TextRun({
          text: part, font: 'Arial',
          size:    l.size   ?? 19,
          bold:    l.bold   ?? false,
          color:   l.color  ?? '222222',
          italics: l.italic ?? false,
        })],
      }));
    });
  });
  return new TableCell({
    borders: thinBorder('DDDDDD'),
    width:   { size: w, type: WidthType.DXA },
    shading: { fill, type: ShadingType.CLEAR },
    margins: { top: 70, bottom: 70, left: 150, right: 150 },
    children,
  });
}

function numCell(text, w, isAlt = false, align = AlignmentType.CENTER) {
  let displayText = String(text);
  if (displayText.endsWith(' €')) {
    const raw = displayText.slice(0, -2).replace(',', '.');
    const n   = parseFloat(raw);
    if (!isNaN(n)) displayText = formatNum(n) + ' €';
  }
  return new TableCell({
    borders:       thinBorder('DDDDDD'),
    width:         { size: w, type: WidthType.DXA },
    shading:       { fill: isAlt ? 'EEF3F3' : C.WHITE, type: ShadingType.CLEAR },
    margins:       { top: 70, bottom: 70, left: 150, right: 150 },
    verticalAlign: VerticalAlign.TOP,
    children: [new Paragraph({
      alignment: align,
      spacing:   { before: 0, after: 0 },
      children:  [new TextRun({ text: displayText, font: 'Arial', size: 19, color: '222222' })],
    })],
  });
}

// ── Shared meta table ─────────────────────────────────────────────────────────
function buildMetaTable(rows) {
  return new Table({
    width: { size: 9326, type: WidthType.DXA },
    columnWidths: [1500, 4326, 1500, 2000],
    rows: rows.map(r => new TableRow({ children: r })),
  });
}

module.exports = {
  C, noBorder, thinBorder, para, emptyPara,
  labelCell, valueCell,
  buildHeader, buildSeparator, buildFooter,
  buildTotals,
  headerItemCell, conceptCell, numCell,
  buildMetaTable,
  formatDate, formatNum,
};
