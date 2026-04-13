const {
  Document, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
} = require('docx');

const {
  C, noBorder, thinBorder, para, emptyPara,
  labelCell, valueCell,
  buildHeader, buildSeparator, buildFooter, buildTotals,
  headerItemCell, conceptCell, numCell, buildMetaTable,
  formatDate,
} = require('./helpers');

function buildPresuposto(d) {
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
          children: [new TextRun({ text: 'ORZAMENTO / PRESUPUESTO', font: 'Arial', size: 28, bold: true, color: C.WHITE })],
        })],
      }),
      new TableCell({
        borders: thinBorder(C.DARK), width: { size: 1500, type: WidthType.DXA },
        shading: { fill: C.DARK, type: ShadingType.CLEAR },
        margins: { top: 150, bottom: 150, left: 150, right: 150 },
        children: [
          para('Nº Orzamento:', { size: 16, bold: true, color: C.GRAY_LINE }),
          para(d.num || '—',    { size: 22, bold: true, color: C.WHITE, align: AlignmentType.CENTER }),
        ],
      }),
      new TableCell({
        borders: thinBorder(C.DARK), width: { size: 1500, type: WidthType.DXA },
        shading: { fill: C.DARK, type: ShadingType.CLEAR },
        margins: { top: 150, bottom: 150, left: 150, right: 150 },
        children: [
          para('Data / Fecha:', { size: 16, bold: true, color: C.GRAY_LINE }),
          para(formatDate(d.fecha),  { size: 20, bold: true, color: C.WHITE }),
        ],
      }),
    ]})],
  });

  // ── Meta table ─────────────────────────────────────────────────────────────
  const metaTable = buildMetaTable([
    [labelCell('Cliente:', 1500), valueCell(d.cnome || '—', 4326, { bold: true }),
     labelCell('NIF/CIF:', 1500), valueCell(d.cnif || '—', 2000)],
    [labelCell('Enderezo:', 1500), valueCell(d.cdir || '—', 4326),
     labelCell('C.P.:', 1500), valueCell(d.ccp || '—', 2000)],
    [labelCell('Válido ata:', 1500), valueCell(formatDate(d.validez), 4326),
     new TableCell({ borders: noBorder(), width: { size: 3500, type: WidthType.DXA }, columnSpan: 2,
       margins: { top: 80, bottom: 80, left: 0, right: 0 }, children: [para('')] })],
  ]);

  // ── Object description ─────────────────────────────────────────────────────
  const objPara = new Paragraph({
    spacing: { before: 240, after: 160 },
    children: [
      new TextRun({ text: 'Obxecto / Objeto: ', font: 'Arial', size: 20, bold: true, color: C.DARK }),
      new TextRun({ text: d.objeto || '—', font: 'Arial', size: 20, color: '222222' }),
    ],
  });

  // ── Items table ────────────────────────────────────────────────────────────
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

  // ── Notes ──────────────────────────────────────────────────────────────────
  // Build note lines — use custom notas if provided, else default text
  const notaLines = d.notas && d.notas.trim()
    ? d.notas.trim().split('\n').filter(l => l.trim()).map(l =>
        para('• ' + l.trim(), { size: 17, color: C.GRAY_TEXT })
      )
    : [
        para('• Non se inclúen remates interiores nin pintura / No se incluyen remates interiores ni pintura.', { size: 17, color: C.GRAY_TEXT }),
        para('• Os prezos inclúen man de obra e materiais VELUX / Los precios incluyen mano de obra y materiales VELUX.', { size: 17, color: C.GRAY_TEXT }),
      ];

  const notesTable = new Table({
    width: { size: 9326, type: WidthType.DXA }, columnWidths: [9326],
    rows: [new TableRow({ children: [new TableCell({
      borders: { top: { style: BorderStyle.SINGLE, size: 4, color: C.GRAY_LINE }, bottom: noBorder().bottom, left: noBorder().left, right: noBorder().right },
      shading: { fill: C.LIGHT_BG, type: ShadingType.CLEAR },
      margins: { top: 120, bottom: 120, left: 200, right: 200 },
      children: [
        para('Notas / Notas:', { size: 18, bold: true, color: C.DARK }),
        ...notaLines,
      ],
    })]})],
  });

  // ── Payment ────────────────────────────────────────────────────────────────
  // d.modalidadePago: 'unico' | 'dous'
  let payTable;

  if (d.modalidadePago === 'unico') {
    // Single full-width payment cell
    const totalPago = d.pagoUnico ? d.pagoUnico + ' €' : d.total || '—';
    payTable = new Table({
      width: { size: 9326, type: WidthType.DXA }, columnWidths: [9326],
      rows: [new TableRow({ children: [
        new TableCell({
          borders: thinBorder(C.DARK), shading: { fill: C.LIGHT_BG, type: ShadingType.CLEAR },
          margins: { top: 130, bottom: 130, left: 250, right: 250 }, width: { size: 9326, type: WidthType.DXA },
          children: [
            para('Pago único ao finalizar / Pago único al finalizar', { bold: true, color: C.DARK, size: 19 }),
            para('Abonarase o importe total unha vez rematada a instalación.', { size: 17, color: C.GRAY_TEXT }),
            para('Se abonará el importe total una vez terminada la instalación.', { size: 17, color: C.GRAY_TEXT, italic: true }),
            para(totalPago, { bold: true, size: 24, color: C.DARK }),
          ],
        }),
      ]})],
    });
  } else {
    // Two payments
    const pago1 = d.pago1 ? d.pago1 + ' €' : '—';
    const pago2 = d.pago2 ? d.pago2 + ' €' : '—';
    payTable = new Table({
      width: { size: 9326, type: WidthType.DXA }, columnWidths: [4663, 4663],
      rows: [new TableRow({ children: [
        new TableCell({
          borders: thinBorder(C.DARK), shading: { fill: C.LIGHT_BG, type: ShadingType.CLEAR },
          margins: { top: 120, bottom: 120, left: 200, right: 200 }, width: { size: 4663, type: WidthType.DXA },
          children: [
            para('1ª Factura — Provisión de materiais', { bold: true, color: C.DARK, size: 19 }),
            para('1ª Factura — Provisión de materiales', { size: 17, color: C.GRAY_TEXT, italic: true }),
            para(pago1, { bold: true, size: 22, color: C.DARK }),
            para('(ó confirmar o pedido / al confirmar el pedido)', { size: 16, italic: true, color: C.GRAY_TEXT }),
          ],
        }),
        new TableCell({
          borders: thinBorder(C.DARK), shading: { fill: C.LIGHT_BG, type: ShadingType.CLEAR },
          margins: { top: 120, bottom: 120, left: 200, right: 200 }, width: { size: 4663, type: WidthType.DXA },
          children: [
            para('2ª Factura — Man de obra', { bold: true, color: C.DARK, size: 19 }),
            para('2ª Factura — Mano de obra', { size: 17, color: C.GRAY_TEXT, italic: true }),
            para(pago2, { bold: true, size: 22, color: C.DARK }),
            para('(ó rematar a instalación / al terminar la instalación)', { size: 16, italic: true, color: C.GRAY_TEXT }),
          ],
        }),
      ]})],
    });
  }

  // ── Signature ──────────────────────────────────────────────────────────────
  const sigTable = new Table({
    width: { size: 9326, type: WidthType.DXA }, columnWidths: [4663, 4663],
    rows: [new TableRow({ children: [
      new TableCell({
        borders: thinBorder(C.GRAY_LINE), margins: { top: 200, bottom: 400, left: 200, right: 200 }, width: { size: 4663, type: WidthType.DXA },
        children: [
          para('Conforme o cliente / Conforme el cliente:', { size: 18, bold: true, color: C.DARK }),
          emptyPara(200),
          para('Sinatura / Firma: ____________________________', { size: 17, color: C.GRAY_TEXT }),
          emptyPara(80),
          para('Data / Fecha: _______________________________', { size: 17, color: C.GRAY_TEXT }),
        ],
      }),
      new TableCell({
        borders: thinBorder(C.GRAY_LINE), margins: { top: 200, bottom: 400, left: 200, right: 200 }, width: { size: 4663, type: WidthType.DXA },
        children: [
          para('Por Garrido Fontal SLU:', { size: 18, bold: true, color: C.DARK }),
          emptyPara(200),
          para('Sinatura / Firma: ____________________________', { size: 17, color: C.GRAY_TEXT }),
          emptyPara(80),
          para('Data / Fecha: _______________________________', { size: 17, color: C.GRAY_TEXT }),
        ],
      }),
    ]})],
  });

  return new Document({
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 800, right: 900, bottom: 800, left: 900 } } },
      footers: { default: buildFooter() },
      children: [
        buildHeader(), buildSeparator(), titleTable,
        emptyPara(80), metaTable, objPara,
        emptyPara(60), itemsTable,
        emptyPara(60), notesTable,
        emptyPara(80), buildTotals(d.base, d.ivaPct, d.ivaVal, d.total),
        emptyPara(100),
        para('Forma de pago / Forma de pago:', { size: 20, bold: true, color: C.DARK }),
        emptyPara(80), payTable,
        emptyPara(120), sigTable,
      ],
    }],
  });
}

module.exports = { buildPresuposto };
