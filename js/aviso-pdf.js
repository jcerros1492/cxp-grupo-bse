// Aviso de pago al proveedor (PDF): página 1 con lo pagado y, si existe, el comprobante
// individual del banco unido al final. Nunca se adjunta el comprobante general del débito.
// Usa jsPDF (página del aviso) y pdf-lib (para unir el comprobante).

const ANCHO = 215.9, ALTO = 279.4, M = 16, UTIL = ANCHO - 2 * M;
const NEGRO = [17, 17, 17], GRIS = [85, 85, 85];

async function imagenComoDataURL(url) {
  try {
    const r = await fetch(url);
    const b = await r.blob();
    return await new Promise((ok) => { const fr = new FileReader(); fr.onload = () => ok(fr.result); fr.readAsDataURL(b); });
  } catch { return null; }
}

/**
 * @param d { sociedad, proveedor, fecha, comprobante, documento, simbolo, facturas: [{documento, fechaFactura, monto, pagado, saldo, tipo}],
 *            cargos, total, cuenta, descripcion, generado, fmtNum, adjunto: { bytes: Uint8Array, tipo } | null }
 * @returns Uint8Array del PDF
 */
export async function crearAvisoPago(d) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: 'mm', format: 'letter', orientation: 'portrait' });
  doc.setProperties({ title: `Aviso de pago - ${d.proveedor}`, creator: 'Cuentas por Pagar · Grupo BSE' });
  const n = d.fmtNum;
  const paginas = d.adjunto ? 2 : 1;

  // Encabezado
  const logo = await imagenComoDataURL('img/emblema-oscuro.png');
  if (logo) doc.addImage(logo, 'PNG', M, 14, 12.5, 12);
  doc.setFont('helvetica', 'bold'); doc.setFontSize(17); doc.setTextColor(...NEGRO);
  doc.text('GRUPO BSE', M + 15, 20.5);
  doc.setFont('helvetica', 'normal'); doc.setFontSize(8.5); doc.setTextColor(...GRIS);
  doc.text('Cuentas por Pagar', M + 15, 25);
  doc.setFont('helvetica', 'bold'); doc.setFontSize(17); doc.setTextColor(...NEGRO);
  doc.text('AVISO DE PAGO', ANCHO - M, 19.5, { align: 'right' });
  doc.setFont('helvetica', 'normal'); doc.setFontSize(9); doc.setTextColor(...GRIS);
  doc.text('Fecha de pago', ANCHO - M - 23, 25, { align: 'right' });
  doc.setFont('helvetica', 'bold'); doc.setTextColor(...NEGRO); doc.text(d.fecha, ANCHO - M, 25, { align: 'right' });
  doc.setFont('helvetica', 'normal'); doc.setTextColor(...GRIS);
  doc.text('Comprobante', ANCHO - M - 23, 29.5, { align: 'right' });
  doc.setFont('helvetica', 'bold'); doc.setTextColor(...NEGRO); doc.text(String(d.comprobante || ''), ANCHO - M, 29.5, { align: 'right' });
  doc.setDrawColor(...NEGRO); doc.setLineWidth(0.7); doc.line(M, 33, ANCHO - M, 33);

  // Partes
  let y = 38;
  const caja = (x, w, et, val) => {
    doc.setDrawColor(213, 218, 224); doc.setLineWidth(0.3); doc.roundedRect(x, y, w, 15, 1.5, 1.5);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(8.5); doc.setTextColor(...GRIS); doc.text(et, x + 4, y + 5.5);
    doc.setFont('helvetica', 'bold'); doc.setTextColor(...NEGRO);
    let t = 11; doc.setFontSize(t);
    while (t > 7.5 && doc.getTextWidth(val) > w - 8) doc.setFontSize(t -= 0.5); // nombres largos: letra más pequeña
    doc.text(doc.splitTextToSize(val, w - 8)[0], x + 4, y + 11.5);
  };
  caja(M, UTIL / 2 - 3, 'Pagado por', d.sociedad);
  caja(M + UTIL / 2 + 3, UTIL / 2 - 3, 'Proveedor', d.proveedor);
  y += 23;
  doc.setFont('helvetica', 'normal'); doc.setFontSize(10); doc.setTextColor(...NEGRO);
  doc.text('Estimados señores de ', M, y);
  doc.setFont('helvetica', 'bold'); doc.text(d.proveedor + ':', M + doc.getTextWidth('Estimados señores de '), y);
  doc.setFont('helvetica', 'normal');
  doc.text('Les informamos que realizamos el siguiente pago mediante transferencia bancaria.', M, y + 5);
  y += 10;

  // Facturas
  const cuerpo = d.facturas.map((f) => [String(f.documento), f.fechaFactura, n(f.monto), n(f.pagado), n(f.saldo), f.tipo]);
  if (d.cargos > 0) cuerpo.push(['', '', 'Intereses / otros cargos', n(d.cargos), '', '']);
  const w = [0.15, 0.16, 0.2, 0.2, 0.16, 0.13].map((p) => p * UTIL);
  doc.autoTable({
    startY: y, margin: { left: M, right: M },
    head: [['Factura', 'Fecha de factura', 'Monto de la factura', 'Monto pagado', 'Saldo pendiente', 'Tipo']],
    body: cuerpo,
    foot: [[{ content: `Total pagado · ${d.facturas.length} ${d.facturas.length === 1 ? 'factura' : 'facturas'}`, colSpan: 3, styles: { halign: 'right' } },
      { content: d.simbolo + n(d.total), styles: { halign: 'right' } }, '', '']],
    theme: 'plain',
    styles: { font: 'helvetica', fontSize: 9.5, cellPadding: 2, textColor: NEGRO, halign: 'center', lineColor: [227, 227, 227], lineWidth: { bottom: 0.2 } },
    headStyles: { fontStyle: 'bold', fontSize: 8, textColor: [68, 68, 68], lineColor: [150, 150, 150], lineWidth: { bottom: 0.4 } },
    footStyles: { fontStyle: 'bold', fontSize: 10.5, lineColor: NEGRO, lineWidth: { top: 0.6 } },
    columnStyles: { 0: { cellWidth: w[0] }, 1: { cellWidth: w[1] }, 2: { cellWidth: w[2], halign: 'right' }, 3: { cellWidth: w[3], halign: 'right' }, 4: { cellWidth: w[4], halign: 'right' }, 5: { cellWidth: w[5] } },
  });
  y = doc.lastAutoTable.finalY + 8;

  // Datos de la transferencia
  const datos = [['Comprobante', String(d.comprobante || '')]];
  if (d.documento) datos.push(['N° de documento del banco', String(d.documento)]);
  datos.push(['Fecha de aplicación', d.fecha]);
  if (d.cuenta) datos.push(['Cuenta acreditada', d.cuenta]);
  if (d.descripcion) datos.push(['Descripción en el banco', d.descripcion]);
  const alto = 10 + datos.length * 5.5;
  doc.setFillColor(246, 247, 249); doc.setDrawColor(225, 229, 234); doc.roundedRect(M, y, UTIL, alto, 1.5, 1.5, 'FD');
  doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.setTextColor(...NEGRO); doc.text('Datos de la transferencia', M + 5, y + 6.5);
  datos.forEach(([k, v], i) => {
    doc.setFont('helvetica', 'normal'); doc.setFontSize(9.5); doc.setTextColor(...GRIS); doc.text(k, M + 5, y + 12.5 + i * 5.5);
    doc.setFont('helvetica', k === 'Comprobante' ? 'bold' : 'normal'); doc.setTextColor(...NEGRO); doc.text(v, M + 70, y + 12.5 + i * 5.5);
  });
  y += alto + 9;
  doc.setFont('helvetica', 'normal'); doc.setFontSize(9.5); doc.setTextColor(...NEGRO);
  if (d.adjunto) doc.text('En la página siguiente adjuntamos el comprobante emitido por el banco.', M, y);
  doc.text('Ante cualquier consulta, con gusto les atendemos.', M, y + (d.adjunto ? 5 : 0));

  // Pie
  doc.setDrawColor(221, 221, 221); doc.setLineWidth(0.2); doc.line(M, ALTO - 12, ANCHO - M, ALTO - 12);
  doc.setFontSize(7.5); doc.setTextColor(119, 119, 119);
  doc.text(`Aviso generado por Cuentas por Pagar · Grupo BSE · ${d.generado}`, M, ALTO - 8);
  doc.text(`Página 1 de ${paginas}`, ANCHO - M, ALTO - 8, { align: 'right' });

  const aviso = new Uint8Array(doc.output('arraybuffer'));
  if (!d.adjunto) return aviso;

  // Unir el comprobante individual del banco
  const { PDFDocument, StandardFonts, rgb } = window.PDFLib;
  const final = await PDFDocument.load(aviso);
  const fuente = await final.embedFont(StandardFonts.Helvetica);
  const pieDe = (pag, i, total) => {
    const { width } = pag.getSize();
    const t = `Aviso de pago · ${d.sociedad} -> ${d.proveedor} · comprobante del banco · Página ${i} de ${total}`;
    pag.drawText(t, { x: width / 2 - fuente.widthOfTextAtSize(t, 7) / 2, y: 14, size: 7, font: fuente, color: rgb(0.47, 0.47, 0.47) });
  };
  if (d.adjunto.tipo === 'application/pdf') {
    const otro = await PDFDocument.load(d.adjunto.bytes, { ignoreEncryption: true });
    const copias = await final.copyPages(otro, otro.getPageIndices());
    copias.forEach((p) => final.addPage(p));
  } else {
    const img = d.adjunto.tipo === 'image/png' ? await final.embedPng(d.adjunto.bytes) : await final.embedJpg(d.adjunto.bytes);
    const pag = final.addPage([612, 792]);
    const esc = Math.min((612 - 72) / img.width, (792 - 110) / img.height);
    pag.drawText('Comprobante del banco', { x: 36, y: 750, size: 14, font: await final.embedFont(StandardFonts.HelveticaBold) });
    pag.drawImage(img, { x: 36, y: 740 - img.height * esc - 8, width: img.width * esc, height: img.height * esc });
  }
  const total = final.getPageCount();
  // Corrige "Página 1 de N" de la primera hoja si el comprobante trae varias páginas
  if (total > 2) {
    const p1 = final.getPage(0);
    p1.drawRectangle({ x: 612 - 46 - 80, y: 792 - (ALTO - 6) / 25.4 * 72, width: 82, height: 12, color: rgb(1, 1, 1) });
    const t = `Página 1 de ${total}`;
    p1.drawText(t, { x: 612 - 46 - fuente.widthOfTextAtSize(t, 7.5), y: 792 - (ALTO - 8) / 25.4 * 72 - 2, size: 7.5, font: fuente, color: rgb(0.47, 0.47, 0.47) });
  }
  for (let i = 1; i < total; i++) pieDe(final.getPage(i), i + 1, total);
  return await final.save();
}
