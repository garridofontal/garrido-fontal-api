const {
  Document, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
} = require('docx');

const {
  C, noBorder, thinBorder, para, emptyPara,
  labelCell, valueCell,
  buildHeader, buildSeparator, buildFooter, buildTotals,
  headerItemCell, conceptCell, numCell, buildMetaTable,
  formatCurrency, formatDate,
} = require('./helpers');

function buildFactura(d) {
  // ââ Title bar ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
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
          para('NÂº Factura:', { size: 16, bold: true, color: C.GRAY_LINE }),
          para(d.num || 'â', { size: 22, bold: true, color: C.WHITE, align: AlignmentType.CENTER }),
        ],
      }),
      new TableCell({
        borders: thinBorder(C.DARK), width: { size: 1500, type: WidthType.DXA },
        shading: { fill: C.DARK, type: ShadingType.CLEAR },
        margins: { top: 150, bottom: 150, left: 150, right: 150 },
        children: [
          para('Fecha:', { size: 16, bold: true, color: C.GRAY_LINE }),
          para(formatDate(d.fecha) || 'â', { size: 20, bold: true, color: C.WHITE }),
        ],
      }),
    ]})],
  });

  // ââ Meta table âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
  const metaTable = buildMetaTable([
    [labelCell('Cliente:', 1500), valueCell(d.cnome || 'â', 4326, { bold: true }),
     labelCell('NIF/CIF:', 1500), valueCell(d.cnif || 'â', 2000)],
    [labelCell('Enderezo:', 1500), valueCell(d.cdir || 'â', 4326),
     labelCell('C.P. / Localidade:', 1500), valueCell(d.ccp || 'â', 2000)],
  ]);

  // ââ Items table ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
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
      numCell(l.precio ? formatCurrency(l.precio) : 'â', COLS[2], isAlt, AlignmentType.RIGHT),
      numCell(l.subtotal ? formatCurrency(l.subtotal) : 'â', COLS[3], isAlt, AlignmentType.RIGHT),
    ]});
  });

  const itemsTable = new Table({
    width: { size: 9326, type: WidthType.DXA },
    columnWidths: COLS,
    rows: [
      new TableRow({ children: [
        headerItemCell('Ud.', COLS[0]),
        headerItemCell('Concepto / DescripciÃ³n', COLS[1], AlignmentType.LEFT),
        headerItemCell('Precio Ud.', COLS[2]),
        headerItemCell('Total\nc/IVA incl.', COLS[3]),
      ]}),
      ...itemRows,
    ],
  });

  // ââ Notes (optional) âââââââââââââââââââââââââââââââââââââââââââââââââââââââ
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

  // ââ Bank âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
  const bankTable = new Table({
    width: { size: 9326, type: WidthType.DXA }, columnWidths: [9326],
    rows: [new TableRow({ children: [new TableCell({
      borders: thinBorder(C.DARK),
      shading: { fill: C.LIGHT_BG, type: ShadingType.CLEAR },
      margins: { top: 130, bottom: 130, left: 250, right: 250 },
      children: [
        new Paragraph({
          spacing: { before: 0, after: 60 },
          children: [
            new TextRun({ text: 'Forma de pago / Forma de pago:  ', font: 'Arial', size: 18, bold: true, color: C.DARK }),
            new TextRun({ text: 'Transferencia bancaria ou Adeudo en conta / Transferencia bancaria o Adeudo en cuenta', font: 'Arial', size: 17, color: C.GRAY_TEXT }),
          ],
        }),
        new Paragraph({
          spacing: { before: 0, after: 0 },
          children: [
            new TextRun({ text: `${d.banco || ''}   `, font: 'Arial', size: 18, bold: true, color: C.DARK }),
            new TextRun({ text: d.iban || 'â', font: 'Arial', size: 20, bold: true, color: '222222' }),
          ],
        }),
      ],
    })]})],
  });

  // ââ Registro mercantil âââââââââââââââââââââââââââââââââââââââââââââââââââââ
  const registroPara = new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 160, after: 0 },
    children: [new TextRun({
      text: 'Inscrita R.M. Lugo Â· Tomo 223 Â· Folio 37 Â· Hoja LU4553 Â· InscripciÃ³n 1Âª',
      font: 'Arial', size: 16, italic: true, color: C.GRAY_TEXT,
    })],
  });

  return new Document({
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 600, right: 800, bottom: 600, left: 800 } } },
      footers: { default: buildFooter() },
      children: [
        buildHeader(), buildSeparator(), titleTable,
        emptyPara(80), metaTable,
        emptyPara(140), itemsTable,
        emptyPara(60), ...notesBlock,
        buildTotals(d.base, d.ivaPct, d.ivaVal, d.total),
        emptyPara(140), bankTable,
        registroPara,
      ],
    }],
  });
}

module.exports = { buildFactura };
