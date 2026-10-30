// Genera el PDF del reporte de cuentas por pagar (hoja carta), listo para descargar o compartir.
// Usa jsPDF + autotable (js/vendor). No lleva encabezados ni pies del navegador.

const MM = { izq: 14, der: 14, arriba: 14, abajo: 14 };
const ANCHO = 215.9, ALTO = 279.4;
const UTIL = ANCHO - MM.izq - MM.der;
// Mismas proporciones que el reporte en pantalla: N°, F. factura, F. vencimiento, Días, Saldo, Moneda
const COLS = [0.16, 0.16, 0.17, 0.11, 0.24, 0.16].map((p) => p * UTIL);
const NEGRO = [17, 17, 17], GRIS = [85, 85, 85], ROJO = [180, 35, 24], VERDE = [23, 100, 58], AMBAR = [138, 75, 0];

/**
 * @param d { grupos, porProveedor, rango, moneda, emitido, nombreSociedad(clave), nombreCorto(clave), fmtNum(n), fmtFecha(iso), dias(f), saldo(f), total }
 */
export function crearPdfReporte(d) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: 'mm', format: 'letter', orientation: 'portrait' });
  doc.setProperties({ title: d.titulo || 'Reporte de Cuentas por Pagar', creator: 'Cuentas por Pagar · Grupo BSE' });
  let y = MM.arriba;
  const espacio = (alto) => { if (y + alto > ALTO - MM.abajo) { doc.addPage(); y = MM.arriba; } };

  const encabezado = (soc) => {
    doc.setFont('helvetica', 'bold'); doc.setFontSize(16); doc.setTextColor(...NEGRO);
    doc.text(d.nombreSociedad(soc), MM.izq, y + 5);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(7.5); doc.setTextColor(...GRIS);
    doc.text('Emitido', ANCHO - MM.der, y + 1.5, { align: 'right' });
    doc.setFont('helvetica', 'bold'); doc.setFontSize(9); doc.setTextColor(...NEGRO);
    doc.text(d.emitido.fecha, ANCHO - MM.der, y + 5.5, { align: 'right' });
    doc.setFont('helvetica', 'normal'); doc.setFontSize(7.5); doc.setTextColor(...GRIS);
    doc.text(d.emitido.hora, ANCHO - MM.der, y + 9, { align: 'right' });
    y += 11;
    doc.setFont('helvetica', 'bold'); doc.setFontSize(11); doc.setTextColor(...NEGRO);
    doc.text('Reporte de Cuentas por Pagar', MM.izq, y); y += 5;
    doc.setFont('helvetica', 'normal'); doc.setFontSize(9); doc.setTextColor(...GRIS);
    doc.text(`${d.rango} · ${d.moneda}`, MM.izq, y); y += 4.5;
    doc.setFontSize(7.5);
    doc.setTextColor(...ROJO); doc.text('-', MM.izq + 8, y);
    doc.setTextColor(...GRIS); doc.text('Días:', MM.izq, y); doc.text('vencida hace esos días ·', MM.izq + 11, y);
    doc.setTextColor(...VERDE); doc.text('+', MM.izq + 41, y);
    doc.setTextColor(...GRIS); doc.text('faltan esos días para vencer', MM.izq + 44, y);
    y += 3;
    doc.setDrawColor(...NEGRO); doc.setLineWidth(0.6); doc.line(MM.izq, y, ANCHO - MM.der, y);
    y += 6;
  };
  const encabezadoCorto = (soc) => {
    espacio(30);
    doc.setFont('helvetica', 'bold'); doc.setFontSize(13); doc.setTextColor(...NEGRO);
    doc.text(d.nombreSociedad(soc), MM.izq, y + 4); y += 6.5;
    doc.setDrawColor(...NEGRO); doc.setLineWidth(0.3); doc.line(MM.izq, y, ANCHO - MM.der, y);
    y += 5;
  };
  const filaTotal = (texto, valor, grande) => {
    espacio(12);
    doc.setDrawColor(...NEGRO); doc.setLineWidth(grande ? 0.6 : 0.3); doc.line(MM.izq, y, ANCHO - MM.der, y);
    y += grande ? 6 : 5;
    const xSaldo = MM.izq + COLS[0] + COLS[1] + COLS[2] + COLS[3] + COLS[4] - 9;
    doc.setFont('helvetica', 'bold'); doc.setFontSize(grande ? 11 : 10); doc.setTextColor(...NEGRO);
    doc.text(texto, MM.izq + COLS[0] + COLS[1] + COLS[2] + COLS[3] - 2, y, { align: 'right' });
    doc.text(valor, xSaldo, y, { align: 'right' });
    y += grande ? 7 : 5;
  };

  d.grupos.forEach((g, i) => {
    if (i > 0 && !d.porProveedor) { doc.addPage(); y = MM.arriba; }
    if (i === 0 || !d.porProveedor) encabezado(g.soc); else encabezadoCorto(g.soc);
    for (const x of g.proveedores) {
      espacio(22);
      doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.setTextColor(...NEGRO);
      doc.text(x.prov.nombre, MM.izq, y);
      doc.setLineWidth(0.25); doc.line(MM.izq, y + 0.8, MM.izq + doc.getTextWidth(x.prov.nombre), y + 0.8);
      y += 2;
      const cuerpo = x.facturas.map((f) => {
        const dd = d.dias(f);
        return [String(f.documento), d.fmtFecha(f.fechaFactura), d.fmtFecha(f.fechaVence),
          dd == null ? '' : dd === 0 ? 'Hoy' : (dd > 0 ? '+' : '-') + Math.abs(dd), d.fmtNum(d.saldo(f)), d.moneda];
      });
      const pie = d.porProveedor ? [] : [[
        { content: `Total Cliente · ${x.facturas.length} ${x.facturas.length === 1 ? 'factura' : 'facturas'}`, colSpan: 4, styles: { halign: 'right' } },
        { content: d.fmtNum(x.saldo), styles: { halign: 'right' } }, '',
      ]];
      doc.autoTable({
        startY: y, margin: { left: MM.izq, right: MM.der, top: MM.arriba, bottom: MM.abajo },
        head: [['N° Factura', 'Fecha de\nfactura', 'Fecha de\nvencimiento', 'Días', 'Saldo', 'Moneda']],
        body: cuerpo, foot: pie, theme: 'plain', showFoot: 'lastPage', rowPageBreak: 'avoid',
        styles: { font: 'helvetica', fontSize: 8.5, cellPadding: { top: 1.4, bottom: 1.4, left: 1.5, right: 1.5 }, textColor: NEGRO, halign: 'center', valign: 'middle', lineColor: [227, 227, 227], lineWidth: { bottom: 0.2 } },
        headStyles: { fontStyle: 'bold', fontSize: 7.5, textColor: [68, 68, 68], halign: 'center', valign: 'bottom', lineColor: [150, 150, 150], lineWidth: { bottom: 0.3 } },
        footStyles: { fontStyle: 'bold', fontSize: 9.5, textColor: NEGRO, lineWidth: 0 },
        columnStyles: { 0: { cellWidth: COLS[0] }, 1: { cellWidth: COLS[1] }, 2: { cellWidth: COLS[2] }, 3: { cellWidth: COLS[3], fontStyle: 'bold' }, 4: { cellWidth: COLS[4], halign: 'right', cellPadding: { top: 1.4, bottom: 1.4, left: 1.5, right: 9 } }, 5: { cellWidth: COLS[5] } },
        didParseCell: (h) => {
          if (h.section === 'body' && h.column.index === 3) {
            const t = String(h.cell.raw);
            h.cell.styles.textColor = t === 'Hoy' ? AMBAR : t.startsWith('-') ? ROJO : VERDE;
          }
          if (h.section === 'foot' && h.column.index === 4) h.cell.styles.cellPadding = { top: 1.4, bottom: 1.4, left: 1.5, right: 9 };
        },
      });
      y = doc.lastAutoTable.finalY + 6;
    }
    filaTotal(`Total ${d.nombreCorto(g.soc)} · ${g.cantidad} ${g.cantidad === 1 ? 'factura' : 'facturas'}`, d.fmtNum(g.saldo), true);
    y += 3;
  });
  if (d.total) {
    espacio(16);
    doc.setDrawColor(...NEGRO); doc.setLineWidth(0.6);
    doc.rect(MM.izq, y, UTIL, 11);
    const xSaldo = MM.izq + COLS[0] + COLS[1] + COLS[2] + COLS[3] + COLS[4] - 9;
    doc.setFont('helvetica', 'bold'); doc.setFontSize(11); doc.setTextColor(...NEGRO);
    doc.text(`Total todas las sociedades · ${d.total.n} facturas`, MM.izq + COLS[0] + COLS[1] + COLS[2] + COLS[3] - 2, y + 7, { align: 'right' });
    doc.text(d.fmtNum(d.total.saldo), xSaldo, y + 7, { align: 'right' });
  }
  // Número de página al pie de cada hoja
  const paginas = doc.getNumberOfPages();
  for (let i = 1; i <= paginas; i++) {
    doc.setPage(i);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(...GRIS);
    doc.text(`Página ${i} de ${paginas}`, ANCHO / 2, ALTO - 7, { align: 'center' });
  }
  return doc;
}
