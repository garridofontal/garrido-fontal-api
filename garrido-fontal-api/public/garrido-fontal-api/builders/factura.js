const {
  Document, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
} = require('docx');

const {
  C, noBorder, thinBorder, para, emptyPara,
  labelCell, valueCell,
  buildHeader, buildSeparator, buildFooter, buildTotals,
  headerItemCell, conceptCell, numCell, buildMetaTable,
} = require('./helpers');

function buildFactura(d) {
  // ── Title bar ──────────────────────────────────────────────────────────────
  const titleTable = new Table({
    width: { size: 9326, type: WidthType.DXA },
    columnWidths: [6326, 1500, 1500],
    rows: [new TableRow({ children: [
      new TableCell({
        borders: thinBorder(C.DARK), width: { size: 6326, type: WidthType.DXA },
        shading: { fill: C.DARK, type: ShadingType.CLEAR },
        margins: { top: 150, bottom: 150, left: 300, right: 300 },
        children: [new Paragraph({
          alignment: AlignmentType.LEFT, spacing: { before: 0, after: 0 },
          children: [new TextRun({ text: 'FACTURA / FACTURA', font: 'Arial', size: 28, bold: true, color: C.WHITE })],
        })],
      }),
      new TableCell({
        borders: thinBorder(C.DARK), width: { size: 1500, type: WidthType.DXA },
        shading: { fill: C.DARK, type: ShadingType.CLEAR },
        margins: { top: 150, bottom: 150, left: 150, right: 150 },
        children: [
          para('Nº Factura:', { size: 16, bold: true, color: C.GRAY_LINE }),
          para(d.num || '—', { size: 22, bold: true, color: C.WHITE, align: AlignmentType.CENTER }),
        ],
      }),
      new TableCell({
        borders: thinBorder(C.DARK), width: { size: 1500, type: WidthType.DXA },
        shading: { fill: C.DARK, type: ShadingType.CLEAR },
        margins: { top: 150, bottom: 150, left: 150, right: 150 },
        children: [
          para('Data / Fecha:', { size: 16, bold: true, color: C.GRAY_LINE }),
          para(d.fecha || '—', { size: 20, bold: true, color: C.WHITE }),
        ],
      }),
    ]})],
  });

  // ── Meta table ─────────────────────────────────────────────────────────────
  const metaTable = buildMetaTable([
    [labelCell('Cliente:', 1500), valueCell(d.cnome || '—', 4326, { bold: true }),
     labelCell('NIF/CIF:', 1500), valueCell(d.cnif || '—', 2000)],
    [labelCell('Enderezo:', 1500), valueCell(d.cdir || '—', 4326),
     labelCell('C.P. / Localidade:', 1500), valueCell(d.ccp || '—', 2000)],
  ]);

  // ── Items table ────────────────────────────────────────────────────────────
  const COLS = [600, 5126, 1800, 1800];
  const itemRows = (d.lineas || []).map((l, i) => {
    const isAlt = i % 2 !== 0;
    const conceptLines = [];
    const lines = String(l.concepto || '').split('\n');
    lines.forEach((ln, li) => {
      conceptLines.push({ text: ln, bold: li === 0, color: li === 0 ? '222222' : '888888', italic: li > 0, size: li === 0 ? 19 : 17 });
    });
    return new TableRow({ children: [
      numCell(l.ud || '1', COLS[0], isAlt),
      conceptCell(conceptLines, COLS[1], isAlt),
      numCell(l.precio ? l.precio + ' €' : '—', COLS[2], isAlt, AlignmentType.RIGHT),
      numCell(l.subtotal ? l.subtotal + ' €' : '—', COLS[3], isAlt, AlignmentType.RIGHT),
    ]});
  });

  const itemsTable = new Table({
    width: { size: 9326, type: WidthType.DXA },
    columnWidths: COLS,
    rows: [
      new TableRow({ children: [
        headerItemCell('Ud.', COLS[0]),
        headerItemCell('Concepto / Descripción', COLS[1], AlignmentType.LEFT),
        headerItemCell('Prezo Ud.\nPrecio Ud.', COLS[2]),
        headerItemCell('Total\nc/IVA incl.', COLS[3]),
      ]}),
      ...itemRows,
    ],
  });

  // ── Notes (optional) ───────────────────────────────────────────────────────
  const notesBlock = d.notas ? [
    new Table({
      width: { size: 9326, type: WidthType.DXA }, columnWidths: [9326],
      rows: [new TableRow({ children: [new TableCell({
        borders: { top: { style: BorderStyle.SINGLE, size: 4, color: C.GRAY_LINE }, bottom: noBorder().bottom, left: noBorder().left, right: noBorder().right },
        shading: { fill: C.LIGHT_BG, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 200, right: 200 },
        children: [para(d.notas, { size: 17, color: C.GRAY_TEXT })],
      })]})],
    }),
    emptyPara(120),
  ] : [emptyPara(120)];

  // ── Bank ───────────────────────────────────────────────────────────────────
  const bankTable = new Table({
    width: { size: 9326, type: WidthType.DXA }, columnWidths: [9326],
    rows: [new TableRow({ children: [new TableCell({
      borders: thinBorder(C.DARK),
      shading: { fill: C.LIGHT_BG, type: ShadingType.CLEAR },
      margins: { top: 130, bottom: 130, left: 250, right: 250 },
      children: [
        new Paragraph({
          spacing: { before: 0, after: 40 },
          children: [
            new TextRun({ text: 'Forma de pago / Forma de pago:  ', font: 'Arial', size: 18, bold: true, color: C.DARK }),
            new TextRun({ text: 'Transferencia bancaria', font: 'Arial', size: 18, color: C.GRAY_TEXT }),
          ],
        }),
        new Paragraph({
          spacing: { before: 0, after: 0 },
          children: [
            new TextRun({ text: `${d.banco || ''}   `, font: 'Arial', size: 18, bold: true, color: C.DARK }),
            new TextRun({ text: d.iban || '—', font: 'Arial', size: 20, bold: true, color: '222222' }),
          ],
        }),
      ],
    })]})],
  });

  // ── Registro mercantil ─────────────────────────────────────────────────────
  const registroPara = new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 160, after: 0 },
    children: [new TextRun({
      text: 'Inscrita R.M. Lugo · Tomo 223 · Folio 37 · Hoja LU4553 · Inscripción 1ª',
      font: 'Arial', size: 16, italic: true, color: C.GRAY_TEXT,
    })],
  });

  return new Document({
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, right: 1000, bottom: 1000, left: 1000 } } },
      footers: { default: buildFooter() },
      children: [
        buildHeader(), buildSeparator(), titleTable,
        emptyPara(120), metaTable,
        emptyPara(200), itemsTable,
        emptyPara(80), ...notesBlock,
        buildTotals(d.base, d.ivaPct, d.ivaVal, d.total),
        emptyPara(200), bankTable,
        registroPara,
      ],
    }],
  });
}

module.exports = { buildFactura };
