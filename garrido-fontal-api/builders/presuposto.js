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

function buildPresuposto(d) {
  // ── Title bar
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
          children: [new TextRun({ text: 'PRESUPUESTO', font: 'Arial', size: 28, bold: true, color: C.WHITE })],
        })],
      }),
      new TableCell({
        borders: thinBorder(C.DARK), width: { size: 1500, type: WidthType.DXA },
        shading: { fill: C.DARK, type: ShadingType.CLEAR },
        margins: { top: 150, bottom: 150, left: 150, right: 150 },
        children: [
          para('Nº Presupuesto:', { size: 16, bold: true, color: C.GRAY_LINE }),
          para(d.num || '—', { size: 22, bold: true, color: C.WHITE, align: AlignmentType.CENTER }),
        ],
      }),
      new TableCell({
        borders: thinBorder(C.DARK), width: { size: 1500, type: WidthType.DXA },
        shading: { fill: C.DARK, type: ShadingType.CLEAR },
        margins: { top: 150, bottom: 150, left: 150, right: 150 },
        children: [
          para('Fecha:', { size: 16, bold: true, color: C.GRAY_LINE }),
          para(formatDate(d.fecha) || '—', { size: 20, bold: true, color: C.WHITE }),
        ],
      }),
    ]})],
  });

  // ── Meta table
  const metaTable = buildMetaTable([
    [labelCell('Cliente:', 1500), valueCell(d.cnome || '—', 4326, { bold: true }),
     labelCell('NIF/CIF:', 1500), valueCell(d.cnif || '—', 2000)],
    [labelCell('Dirección:', 1500), valueCell(d.cdir || '—', 4326),
     labelCell('C.P.:', 1500), valueCell(d.ccp || '—', 2000)],
    [labelCell('Válido hasta:', 1500), valueCell(formatDate(d.validez) || '—', 4326),
     new TableCell({ borders: noBorder(), width: { size: 3500, type: WidthType.DXA }, columnSpan: 2,
       margins: { top: 80, bottom: 80, left: 0, right: 0 }, children: [para('')] })],
  ]);

  // ── Object description
  const objPara = new Paragraph({
    spacing: { before: 240, after: 160 },
    children: [
      new TextRun({ text: 'Objeto: ', font: 'Arial', size: 20, bold: true, color: C.DARK }),
      new TextRun({ text: d.objeto || '—', font: 'Arial', size: 20, color: '222222' }),
    ],
  });

  // ── Items table
  const COLS = [600, 5726, 1500, 1500];
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
      numCell(l.precio ? formatCurrency(l.precio) : '—', COLS[2], isAlt, AlignmentType.RIGHT),
      numCell(l.subtotal ? formatCurrency(l.subtotal) : '—', COLS[3], isAlt, AlignmentType.RIGHT),
    ]});
  });

  const itemsTable = new Table({
    width: { size: 9326, type: WidthType.DXA },
    columnWidths: COLS,
    rows: [
      new TableRow({ children: [
        headerItemCell('Ud.', COLS[0]),
        headerItemCell('Concepto / Descripción', COLS[1], AlignmentType.LEFT),
        headerItemCell('Precio Ud.', COLS[2]),
        headerItemCell('Total s/IVA', COLS[3]),
      ]}),
      ...itemRows,
    ],
  });

  // ── Notes
  const notesTable = new Table({
    width: { size: 9326, type: WidthType.DXA }, columnWidths: [9326],
    rows: [new TableRow({ children: [new TableCell({
      borders: { top: { style: BorderStyle.SINGLE, size: 4, color: C.GRAY_LINE }, bottom: noBorder().bottom, left: noBorder().left, right: noBorder().right },
      shading: { fill: C.LIGHT_BG, type: ShadingType.CLEAR },
      margins: { top: 120, bottom: 120, left: 200, right: 200 },
      children: [
        para('Notas:', { size: 18, bold: true, color: C.DARK }),
        para('• No se incluyen remates interiores ni pintura.', { size: 17, color: C.GRAY_TEXT }),
        para('• Los precios incluyen mano de obra y materiales VELUX.', { size: 17, color: C.GRAY_TEXT }),
      ],
    })]})],
  });

  // ── Payment
  const buildPayCell = (width, title, importe, legend) => new TableCell({
    borders: thinBorder(C.DARK),
    shading: { fill: C.LIGHT_BG, type: ShadingType.CLEAR },
    margins: { top: 100, bottom: 100, left: 180, right: 180 },
    width: { size: width, type: WidthType.DXA },
    children: [
      para(title,   { bold: true, color: C.DARK, size: 18 }),
      para(importe, { bold: true, size: 20, color: C.DARK }),
      para('(' + legend + ')', { size: 15, italic: true, color: C.GRAY_TEXT }),
    ],
  });

  let payTable;
  if (d.modalidadePago === 'unico') {
    const totalPago = d.pagoUnico ? formatCurrency(d.pagoUnico) : formatCurrency(d.total);
    payTable = new Table({
      width: { size: 9326, type: WidthType.DXA }, columnWidths: [9326],
      rows: [new TableRow({ children: [
        new TableCell({
          borders: thinBorder(C.DARK), shading: { fill: C.LIGHT_BG, type: ShadingType.CLEAR },
          margins: { top: 110, bottom: 110, left: 250, right: 250 }, width: { size: 9326, type: WidthType.DXA },
          children: [
            para('Pago único al finalizar', { bold: true, color: C.DARK, size: 19 }),
            para('Se abonará el importe total una vez terminada la instalación.', { size: 16, color: C.GRAY_TEXT }),
            para(totalPago, { bold: true, size: 22, color: C.DARK }),
          ],
        }),
      ]})],
    });
  } else if (d.modalidadePago === 'tres') {
    const pago1 = d.pago1 ? formatCurrency(d.pago1) : '—';
    const pago2 = d.pago2 ? formatCurrency(d.pago2) : '—';
    const pago3 = d.pago3 ? formatCurrency(d.pago3) : '—';
    payTable = new Table({
      width: { size: 9326, type: WidthType.DXA }, columnWidths: [3109, 3109, 3108],
      rows: [new TableRow({ children: [
        buildPayCell(3109, '1ª — Provisión de materiales', pago1, 'al confirmar el pedido'),
        buildPayCell(3109, '2ª — Inicio de la obra',       pago2, 'al comenzar la instalación'),
        buildPayCell(3108, '3ª — Mano de obra final',      pago3, 'al terminar la instalación'),
      ]})],
    });
  } else {
    const pago1 = d.pago1 ? formatCurrency(d.pago1) : '—';
    const pago2 = d.pago2 ? formatCurrency(d.pago2) : '—';
    payTable = new Table({
      width: { size: 9326, type: WidthType.DXA }, columnWidths: [4663, 4663],
      rows: [new TableRow({ children: [
        buildPayCell(4663, '1ª Factura — Provisión de materiales', pago1, 'al confirmar el pedido'),
        buildPayCell(4663, '2ª Factura — Mano de obra',            pago2, 'al terminar la instalación'),
      ]})],
    });
  }

  // ── Signature
  const sigTable = new Table({
    width: { size: 9326, type: WidthType.DXA }, columnWidths: [4663, 4663],
    rows: [new TableRow({ children: [
      new TableCell({
        borders: thinBorder(C.GRAY_LINE), margins: { top: 150, bottom: 280, left: 200, right: 200 }, width: { size: 4663, type: WidthType.DXA },
        children: [
          para('Conforme el cliente:', { size: 17, bold: true, color: C.DARK }),
          emptyPara(140),
          para('Firma: ____________________________', { size: 16, color: C.GRAY_TEXT }),
          emptyPara(50),
          para('Fecha: _______________________________', { size: 16, color: C.GRAY_TEXT }),
        ],
      }),
      new TableCell({
        borders: thinBorder(C.GRAY_LINE), margins: { top: 150, bottom: 280, left: 200, right: 200 }, width: { size: 4663, type: WidthType.DXA },
        children: [
          para('Por Garrido Fontal SLU:', { size: 17, bold: true, color: C.DARK }),
          emptyPara(140),
          para('Firma: ____________________________', { size: 16, color: C.GRAY_TEXT }),
          emptyPara(50),
          para('Fecha: _______________________________', { size: 16, color: C.GRAY_TEXT }),
        ],
      }),
    ]})],
  });

  return new Document({
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 600, right: 800, bottom: 600, left: 800 } } },
      footers: { default: buildFooter() },
      children: [
        buildHeader(), buildSeparator(), titleTable,
        emptyPara(80), metaTable, objPara,
        emptyPara(40), itemsTable,
        emptyPara(60), notesTable,
        emptyPara(80), buildTotals(d.base, d.ivaPct, d.ivaVal, d.total),
        emptyPara(120),
        para('Forma de pago:', { size: 19, bold: true, color: C.DARK }),
        emptyPara(60), payTable,
        emptyPara(140), sigTable,
      ],
    }],
  });
}

module.exports = { buildPresuposto };
