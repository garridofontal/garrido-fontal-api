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
          children: [new TextRun({ text: 'PRESUPUESTO', font: 'Arial', size: 28, bold: true, color: C.WHITE })],
        })],
      }),
      new TableCell({
        borders: thinBorder(C.DARK), width: { size: 1500, type: WidthType.DXA },
        shading: { fill: C.DARK, type: ShadingType.CLEAR },
        margins: { top: 150, bottom: 150, left: 150, right: 150 },
        children: [
          para('NÂº Orzamento:', { size: 16, bold: true, color: C.GRAY_LINE }),
          para(d.num || 'â',    { size: 22, bold: true, color: C.WHITE, align: AlignmentType.CENTER }),
        ],
      }),
      new TableCell({
        borders: thinBorder(C.DARK), width: { size: 1500, type: WidthType.DXA },
        shading: { fill: C.DARK, type: ShadingType.CLEAR },
        margins: { top: 150, bottom: 150, left: 150, right: 150 },
        children: [
          para('Fecha:', { size: 16, bold: true, color: C.GRAY_LINE }),
          para(formatDate(d.fecha) || 'â',  { size: 20, bold: true, color: C.WHITE }),
        ],
      }),
    ]})],
  });

  // ââ Meta table âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
  const metaTable = buildMetaTable([
    [labelCell('Cliente:', 1500), valueCell(d.cnome || 'â', 4326, { bold: true }),
     labelCell('NIF/CIF:', 1500), valueCell(d.cnif || 'â', 2000)],
    [labelCell('Dirección:', 1500), valueCell(d.cdir || 'â', 4326),
     labelCell('C.P.:', 1500), valueCell(d.ccp || 'â', 2000)],
    [labelCell('VÃ¡lido ata:', 1500), valueCell(formatDate(d.validez) || 'â', 4326),
     new TableCell({ borders: noBorder(), width: { size: 3500, type: WidthType.DXA }, columnSpan: 2,
       margins: { top: 80, bottom: 80, left: 0, right: 0 }, children: [para('')] })],
  ]);

  // ââ Object description âââââââââââââââââââââââââââââââââââââââââââââââââââââ
  const objPara = new Paragraph({
    spacing: { before: 240, after: 160 },
    children: [
      new TextRun({ text: 'Objeto: ', font: 'Arial', size: 20, bold: true, color: C.DARK }),
      new TextRun({ text: d.objeto || 'â', font: 'Arial', size: 20, color: '222222' }),
    ],
  });

  // ââ Items table ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
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
        headerItemCell('Total\ns/IVA', COLS[3]),
      ]}),
      ...itemRows,
    ],
  });

  // ââ Notes ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
  const notesTable = new Table({
    width: { size: 9326, type: WidthType.DXA }, columnWidths: [9326],
    rows: [new TableRow({ children: [new TableCell({
      borders: { top: { style: BorderStyle.SINGLE, size: 4, color: C.GRAY_LINE }, bottom: noBorder().bottom, left: noBorder().left, right: noBorder().right },
      shading: { fill: C.LIGHT_BG, type: ShadingType.CLEAR },
      margins: { top: 120, bottom: 120, left: 200, right: 200 },
      children: [
        para('Notas:', { size: 18, bold: true, color: C.DARK }),
        para('â¢ Non se inclÃºen remates interiores nin pintura / No se incluyen remates interiores ni pintura.', { size: 17, color: C.GRAY_TEXT }),
        para('â¢ Os prezos inclÃºen man de obra e materiais VELUX / Los precios incluyen mano de obra y materiales VELUX.', { size: 17, color: C.GRAY_TEXT }),
      ],
    })]})],
  });

  // ââ Payment ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
  // d.modalidadePago: 'unico' | 'dous' | 'tres'
  const buildPayCell = (width, titleGl, titleEs, importe, legendGl, legendEs) => new TableCell({
    borders: thinBorder(C.DARK),
    shading: { fill: C.LIGHT_BG, type: ShadingType.CLEAR },
    margins: { top: 100, bottom: 100, left: 180, right: 180 },
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
    const totalPago = d.pagoUnico ? formatCurrency(d.pagoUnico) : formatCurrency(d.total);
    payTable = new Table({
      width: { size: 9326, type: WidthType.DXA }, columnWidths: [9326],
      rows: [new TableRow({ children: [
        new TableCell({
          borders: thinBorder(C.DARK), shading: { fill: C.LIGHT_BG, type: ShadingType.CLEAR },
          margins: { top: 110, bottom: 110, left: 250, right: 250 }, width: { size: 9326, type: WidthType.DXA },
          children: [
            para('Pago Ãºnico ao finalizar / Pago Ãºnico al finalizar', { bold: true, color: C.DARK, size: 19 }),
            para('Abonarase o importe total unha vez rematada a instalaciÃ³n.', { size: 16, color: C.GRAY_TEXT }),
            para('Se abonarÃ¡ el importe total una vez terminada la instalaciÃ³n.', { size: 16, color: C.GRAY_TEXT, italic: true }),
            para(totalPago, { bold: true, size: 22, color: C.DARK }),
          ],
        }),
      ]})],
    });
  } else if (d.modalidadePago === 'tres') {
    const pago1 = d.pago1 ? formatCurrency(d.pago1) : 'â';
    const pago2 = d.pago2 ? formatCurrency(d.pago2) : 'â';
    const pago3 = d.pago3 ? formatCurrency(d.pago3) : 'â';
    payTable = new Table({
      width: { size: 9326, type: WidthType.DXA }, columnWidths: [3109, 3109, 3108],
      rows: [new TableRow({ children: [
        buildPayCell(3109, '1Âª â ProvisiÃ³n de materiais', '1Âª â ProvisiÃ³n de materiales', pago1,
          'Ã³ confirmar o pedido', 'al confirmar el pedido'),
        buildPayCell(3109, '2Âª â Inicio da obra', '2Âª â Inicio de la obra', pago2,
          'Ã³ comezar a instalaciÃ³n', 'al comenzar la instalaciÃ³n'),
        buildPayCell(3108, '3Âª â Man de obra final', '3Âª â Mano de obra final', pago3,
          'Ã³ rematar a instalaciÃ³n', 'al terminar la instalaciÃ³n'),
      ]})],
    });
  } else {
    const pago1 = d.pago1 ? formatCurrency(d.pago1) : 'â';
    const pago2 = d.pago2 ? formatCurrency(d.pago2) : 'â';
    payTable = new Table({
      width: { size: 9326, type: WidthType.DXA }, columnWidths: [4663, 4663],
      rows: [new TableRow({ children: [
        buildPayCell(4663, '1Âª Factura â ProvisiÃ³n de materiais', '1Âª Factura â ProvisiÃ³n de materiales', pago1,
          'Ã³ confirmar o pedido', 'al confirmar el pedido'),
        buildPayCell(4663, '2Âª Factura â Man de obra', '2Âª Factura â Mano de obra', pago2,
          'Ã³ rematar a instalaciÃ³n', 'al terminar la instalaciÃ³n'),
      ]})],
    });
  }

  // ââ Signature ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
  const sigTable = new Table({
    width: { size: 9326, type: WidthType.DXA }, columnWidths: [4663, 4663],
    rows: [new TableRow({ children: [
      new TableCell({
        borders: thinBorder(C.GRAY_LINE), margins: { top: 150, bottom: 280, left: 200, right: 200 }, width: { size: 4663, type: WidthType.DXA },
        children: [
          para('Conforme el cliente:', { size: 17, bold: true, color: C.DARK }),
          emptyPara(140),
          para('Sinatura / Firma: ____________________________', { size: 16, color: C.GRAY_TEXT }),
          emptyPara(50),
          para('Data / Fecha: _______________________________', { size: 16, color: C.GRAY_TEXT }),
        ],
      }),
      new TableCell({
        borders: thinBorder(C.GRAY_LINE), margins: { top: 150, bottom: 280, left: 200, right: 200 }, width: { size: 4663, type: WidthType.DXA },
        children: [
          para('Por Garrido Fontal SLU:', { size: 17, bold: true, color: C.DARK }),
          emptyPara(140),
          para('Sinatura / Firma: ____________________________', { size: 16, color: C.GRAY_TEXT }),
          emptyPara(50),
          para('Data / Fecha: _______________________________', { size: 16, color: C.GRAY_TEXT }),
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
