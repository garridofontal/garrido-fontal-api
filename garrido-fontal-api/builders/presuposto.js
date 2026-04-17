const {
  Document, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
} = require('docx');
const {
  C, noBorder, thinBorder, para, emptyPara,
  labelCell, valueCell, buildHeader, buildSeparator, buildFooter,
  buildTotals, headerItemCell, conceptCell, numCell, buildMetaTable, formatDate,
} = require('./helpers');

function buildPresuposto(d) {

  // ── Title bar ─────────────────────────────────────────────────────────────────
  const titleTable = new Table({
    width: { size: 9326, type: WidthType.DXA },
    columnWidths: [6326, 1500, 1500],
    rows: [new TableRow({ children: [
      new TableCell({
        borders: thinBorder(C.DARK),
        width: { size: 6326, type: WidthType.DXA },
        shading: { fill: C.DARK, type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 300, right: 300 },
        children: [new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: { before: 0, after: 0 },
          children: [new TextRun({ text: 'ORZAMENTO / PRESUPUESTO', font: 'Arial', size: 26, bold: true, color: C.WHITE })],
        })],
      }),
      new TableCell({
        borders: thinBorder(C.DARK),
        width: { size: 1500, type: WidthType.DXA },
        shading: { fill: C.DARK, type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 150, right: 150 },
        children: [
          para('Nº Orzamento:', { size: 15, bold: true, color: C.GRAY_LINE }),
          para(d.num || '—', { size: 20, bold: true, color: C.WHITE, align: AlignmentType.CENTER }),
        ],
      }),
      new TableCell({
        borders: thinBorder(C.DARK),
        width: { size: 1500, type: WidthType.DXA },
        shading: { fill: C.DARK, type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 150, right: 150 },
        children: [
          para('Data / Fecha:', { size: 15, bold: true, color: C.GRAY_LINE }),
          para(formatDate(d.fecha), { size: 18, bold: true, color: C.WHITE }),
        ],
      }),
    ]})],
  });

  // ── Meta table ────────────────────────────────────────────────────────────────
  const metaTable = buildMetaTable([
    [labelCell('Cliente:', 1500), valueCell(d.cnome || '—', 4326, { bold: true }), labelCell('NIF/CIF:', 1500), valueCell(d.cnif || '—', 2000)],
    [labelCell('Enderezo:', 1500), valueCell(d.cdir || '—', 4326), labelCell('C.P.:', 1500), valueCell(d.ccp || '—', 2000)],
    [labelCell('Válido ata:', 1500), valueCell(formatDate(d.validez), 4326), new TableCell({
      borders: noBorder(), width: { size: 3500, type: WidthType.DXA }, columnSpan: 2,
      margins: { top: 80, bottom: 80, left: 0, right: 0 }, children: [para('')]
    })],
  ]);

  // ── Object description ────────────────────────────────────────────────────────
  const objPara = new Paragraph({
    spacing: { before: 160, after: 120 },
    children: [
      new TextRun({ text: 'Obxecto / Objeto: ', font: 'Arial', size: 19, bold: true, color: C.DARK }),
      new TextRun({ text: d.objeto || '—', font: 'Arial', size: 19, color: '222222' }),
    ],
  });

  // ── Items table ───────────────────────────────────────────────────────────────
  const COLS = [600, 5726, 1500, 1500];
  const itemRows = (d.lineas || []).map((l, i) => {
    const isAlt = i % 2 !== 0;
    const conceptLines = [];
    const lines = String(l.concepto || '').split('\n');
    lines.forEach((ln, li) => {
      conceptLines.push({
        text: ln,
        bold: li === 0,
        color: li === 0 ? '222222' : '888888',
        italic: li > 0,
        size: li === 0 ? 18 : 16,
      });
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
        headerItemCell('Total\ns/IVA', COLS[3]),
      ]}),
      ...itemRows,
    ],
  });

  // ── Notes ─────────────────────────────────────────────────────────────────────
  const notaLines = d.notas && d.notas.trim()
    ? d.notas.trim().split('\n').filter(l => l.trim()).map(l =>
        para('• ' + l.trim(), { size: 16, color: C.GRAY_TEXT })
      )
    : [
        para('• Non se inclúen remates interiores nin pintura / No se incluyen remates interiores ni pintura.', { size: 16, color: C.GRAY_TEXT }),
        para('• Os prezos inclúen man de obra e materiais VELUX / Los precios incluyen mano de obra y materiales VELUX.', { size: 16, color: C.GRAY_TEXT }),
      ];

  const notesTable = new Table({
    width: { size: 9326, type: WidthType.DXA },
    columnWidths: [9326],
    rows: [new TableRow({ children: [new TableCell({
      borders: { top: { style: BorderStyle.SINGLE, size: 4, color: C.GRAY_LINE }, bottom: noBorder().bottom, left: noBorder().left, right: noBorder().right },
      shading: { fill: C.LIGHT_BG, type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 200, right: 200 },
      children: [
        para('Notas / Notas:', { size: 17, bold: true, color: C.DARK }),
        ...notaLines,
      ],
    })]})],
  });

  // ── Payment ───────────────────────────────────────────────────────────────────
  const buildPayCell = (width, titleGl, titleEs, importe, legendGl, legendEs) => new TableCell({
    borders: thinBorder(C.DARK),
    shading: { fill: C.LIGHT_BG, type: ShadingType.CLEAR },
    margins: { top: 100, bottom: 100, left: 200, right: 200 },
    width: { size: width, type: WidthType.DXA },
    children: [
      para(titleGl, { bold: true, color: C.DARK, size: 18 }),
      para(titleEs, { size: 16, color: C.GRAY_TEXT, italic: true }),
      para(importe, { bold: true, size: 20, color: C.DARK }),
      para(`(${legendGl} / ${legendEs})`, { size: 15, italic: true, color: C.GRAY_TEXT }),
    ],
  });

  let payTable;
  if (d.modalidadePago === 'unico') {
    const totalPago = d.pagoUnico ? d.pagoUnico + ' €' : d.total || '—';
    payTable = new Table({
      width: { size: 9326, type: WidthType.DXA },
      columnWidths: [9326],
      rows: [new TableRow({ children: [
        new TableCell({
          borders: thinBorder(C.DARK),
          shading: { fill: C.LIGHT_BG, type: ShadingType.CLEAR },
          margins: { top: 110, bottom: 110, left: 250, right: 250 },
          width: { size: 9326, type: WidthType.DXA },
          children: [
            para('Pago único ao finalizar / Pago único al finalizar', { bold: true, color: C.DARK, size: 18 }),
            para('Abonarase o importe total unha vez rematada a instalación.', { size: 16, color: C.GRAY_TEXT }),
            para('Se abonará el importe total una vez terminada la instalación.', { size: 16, color: C.GRAY_TEXT, italic: true }),
            para(totalPago, { bold: true, size: 22, color: C.DARK }),
          ],
        }),
      ]})],
    });
  } else if (d.modalidadePago === 'tres') {
    const pago1 = d.pago1 ? d.pago1 + ' €' : '—';
    const pago2 = d.pago2 ? d.pago2 + ' €' : '—';
    const pago3 = d.pago3 ? d.pago3 + ' €' : '—';
    payTable = new Table({
      width: { size: 9326, type: WidthType.DXA },
      columnWidths: [3109, 3109, 3108],
      rows: [new TableRow({ children: [
        buildPayCell(3109, '1ª Factura — Provisión de materiais', '1ª Factura — Provisión de materiales', pago1,
          'ó confirmar o pedido', 'al confirmar el pedido'),
        buildPayCell(3109, '2ª Factura — Inicio da obra', '2ª Factura — Inicio de la obra', pago2,
          'ó comezar a instalación', 'al comenzar la instalación'),
        buildPayCell(3108, '3ª Factura — Man de obra final', '3ª Factura — Mano de obra final', pago3,
          'ó rematar a instalación', 'al terminar la instalación'),
      ]})],
    });
  } else {
    const pago1 = d.pago1 ? d.pago1 + ' €' : '—';
    const pago2 = d.pago2 ? d.pago2 + ' €' : '—';
    payTable = new Table({
      width: { size: 9326, type: WidthType.DXA },
      columnWidths: [4663, 4663],
      rows: [new TableRow({ children: [
        buildPayCell(4663, '1ª Factura — Provisión de materiais', '1ª Factura — Provisión de materiales', pago1,
          'ó confirmar o pedido', 'al confirmar el pedido'),
        buildPayCell(4663, '2ª Factura — Man de obra', '2ª Factura — Mano de obra', pago2,
          'ó rematar a instalación', 'al terminar la instalación'),
      ]})],
    });
  }

  // ── Signature ─────────────────────────────────────────────────────────────────
  const sigTable = new Table({
    width: { size: 9326, type: WidthType.DXA },
    columnWidths: [4663, 4663],
    rows: [new TableRow({ children: [
      new TableCell({
        borders: thinBorder(C.GRAY_LINE),
        margins: { top: 180, bottom: 360, left: 200, right: 200 },
        width: { size: 4663, type: WidthType.DXA },
        children: [
          para('Conforme o cliente / Conforme el cliente:', { size: 17, bold: true, color: C.DARK }),
          emptyPara(160),
          para('Sinatura / Firma: ____________________________', { size: 16, color: C.GRAY_TEXT }),
          emptyPara(60),
          para('Data / Fecha: _______________________________', { size: 16, color: C.GRAY_TEXT }),
        ],
      }),
      new TableCell({
        borders: thinBorder(C.GRAY_LINE),
        margins: { top: 180, bottom: 360, left: 200, right: 200 },
        width: { size: 4663, type: WidthType.DXA },
        children: [
          para('Por Garrido Fontal SLU:', { size: 17, bold: true, color: C.DARK }),
          emptyPara(160),
          para('Sinatura / Firma: ____________________________', { size: 16, color: C.GRAY_TEXT }),
          emptyPara(60),
          para('Data / Fecha: _______________________________', { size: 16, color: C.GRAY_TEXT }),
        ],
      }),
    ]})],
  });

  return new Document({
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 700, right: 800, bottom: 700, left: 800 }
        }
      },
      footers: { default: buildFooter() },
      children: [
        buildHeader(),
        buildSeparator(),
        titleTable,
        emptyPara(60),
        metaTable,
        objPara,
        emptyPara(40),
        itemsTable,
        emptyPara(50),
        notesTable,
        emptyPara(60),
        buildTotals(d.base, d.ivaPct, d.ivaVal, d.total),
        emptyPara(80),
        para('Forma de pago / Forma de pago:', { size: 18, bold: true, color: C.DARK }),
        emptyPara(60),
        payTable,
        emptyPara(100),
        sigTable,
      ],
    }],
  });
}

module.exports = { buildPresuposto };
