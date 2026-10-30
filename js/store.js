// Acceso a los datos para las pantallas. Las reglas viven en logica.js y la
// sincronización con Google Sheets en sync.js.

import * as Sync from './sync.js';
import { round2, diasDesde } from './util.js';

const L = window.Logica;
export const datos = () => Sync.datos();
export const suscribir = Sync.suscribir;

// ---------------------------------------------------------------- Sesión

export const usuarioActual = () => Sync.estado.sesion;
export const esNorman = () => Sync.estado.sesion?.rol === 'aprueba';
export const esAdmin = () => Sync.estado.sesion?.rol === 'admin';
export const hayDatos = () => Object.keys(datos().facturas).length > 0;

// ---------------------------------------------------------------- Consultas

export const listaFacturas = () => Object.values(datos().facturas);
export const factura = (id) => datos().facturas[id];
export const proveedor = (codigo) => datos().proveedores[codigo] || { codigo, nombre: String(codigo), alias: '', ibanCRC: '', ibanUSD: '', plazo: 30 };
export const sociedades = () => datos().sociedades;
export const nombreSociedad = (s) => datos().sociedades[s]?.corto || s;
export const cfg = (k) => L.cfg(datos(), k);

export const movimientosDe = (f) => L.movimientosDe(datos(), f.id);
export const pagado = (f) => L.pagado(datos(), f);
export const saldo = (f) => L.saldo(datos(), f);
export const estado = (f) => L.estado(datos(), f);
export const programacionDe = (f) => L.programacionDe(datos(), f.id);
export const diasVencida = (f) => (f.fechaVence ? diasDesde(f.fechaVence) : null);
export const conSaldo = () => listaFacturas().filter((f) => saldo(f) > 0.004);

export const programacionesPendientes = () => Object.values(datos().programaciones).filter((p) => p.estado === 'programado');
export const lotesAplicados = () => Object.values(datos().lotes).sort((a, b) => String(b.aplicado).localeCompare(String(a.aplicado)));

export function totales(lista) {
  const t = { CRC: 0, USD: 0, n: lista.length };
  for (const f of lista) t[f.moneda] = round2(t[f.moneda] + saldo(f));
  return t;
}

// ---------------------------------------------------------------- Programar (Norman)

/** items: [{facturaId, monto}] → devuelve { error } o { avisos } */
export function programarPagos(items, fechaPago, nota = '') {
  return Sync.ejecutar({
    tipo: 'programar', fechaPago, nota,
    items: items.map((it) => ({ progId: Sync.nuevoId(), facturaId: it.facturaId, monto: round2(it.monto) })),
  });
}

export const quitarProgramacion = (progId, motivo) => Sync.ejecutar({ tipo: 'quitarProgramacion', progId, motivo });

// ---------------------------------------------------------------- Lotes por aplicar

const borradores = (() => { try { return JSON.parse(localStorage.getItem('cxp-borradores') || '{}'); } catch { return {}; } })();

export function guardarBorrador(loteClave, codigo, cambios) {
  const k = loteClave + '|' + codigo;
  borradores[k] = { ...(borradores[k] || {}), ...cambios };
  try { localStorage.setItem('cxp-borradores', JSON.stringify(borradores)); } catch { /* sin almacenamiento */ }
}

/** Agrupa lo programado: sociedad + moneda + fecha; dentro, una línea por proveedor. */
export function lotesPendientes() {
  const mapa = new Map();
  for (const p of programacionesPendientes()) {
    const f = factura(p.facturaId);
    if (!f || f.revision) continue; // en revisión: espera a que la marquen como revisada
    const clave = [f.sociedad, f.moneda, p.fechaPago].join('|');
    if (!mapa.has(clave)) mapa.set(clave, { clave, sociedad: f.sociedad, moneda: f.moneda, fecha: p.fechaPago, lineas: new Map() });
    const lote = mapa.get(clave);
    if (!lote.lineas.has(f.codigo)) lote.lineas.set(f.codigo, { codigo: f.codigo, items: [] });
    lote.lineas.get(f.codigo).items.push({ p, f });
  }
  const lotes = [...mapa.values()].sort((a, b) => a.fecha.localeCompare(b.fecha) || a.sociedad.localeCompare(b.sociedad));
  for (const l of lotes) {
    l.lineas = [...l.lineas.values()].map((ln) => completarLinea(l, ln));
    l.total = round2(l.lineas.reduce((s, x) => s + x.total, 0));
    l.cuentaDebito = datos().sociedades[l.sociedad]?.[l.moneda === 'USD' ? 'cuentaUSD' : 'cuentaCRC'] || '';
  }
  return lotes;
}

function completarLinea(lote, ln) {
  const prov = proveedor(ln.codigo);
  const b = borradores[lote.clave + '|' + ln.codigo] || {};
  const facturas = round2(ln.items.reduce((s, x) => s + x.p.monto, 0));
  const extra = round2(b.extra || 0);
  const docs = ln.items.map((x) => x.f.documento).join(' ');
  const alias = prov.alias || prov.nombre.split(/\s+/).slice(0, 2).join(' ');
  const sugerida = `Pago Fact ${docs}${extra > 0 ? ' mas Intereses' : ''} ${alias}`.replace(/\s+/g, ' ').trim();
  return {
    ...ln, prov, extra, facturas, total: round2(facturas + extra),
    descripcion: b.descripcion ?? sugerida, sugerida,
    iban: prov[lote.moneda === 'USD' ? 'ibanUSD' : 'ibanCRC'] || '',
  };
}

export function aplicarLote(lote, { comprobante, fecha }) {
  const loteId = Sync.nuevoId();
  const r = Sync.ejecutar({
    tipo: 'aplicarLote', loteId, comprobante, fecha, sociedad: lote.sociedad, moneda: lote.moneda,
    progIds: lote.lineas.flatMap((ln) => ln.items.map((x) => x.p.id)),
    extras: lote.lineas.filter((ln) => ln.extra > 0).map((ln) => ({ facturaId: ln.items[0].f.id, monto: ln.extra })),
    lineas: lote.lineas.map((ln) => ({ codigo: ln.codigo, total: ln.total, extra: ln.extra, descripcion: ln.descripcion, documentos: ln.items.map((x) => x.f.documento) })),
  });
  if (!r.error) {
    for (const ln of lote.lineas) delete borradores[lote.clave + '|' + ln.codigo];
    try { localStorage.setItem('cxp-borradores', JSON.stringify(borradores)); } catch { /* sin almacenamiento */ }
  }
  return { ...r, loteId };
}

// ---------------------------------------------------------------- Anotaciones compartidas (1.18)
// Se guardan en la pestaña Config: envío al banco, N° de documento por proveedor y comprobantes adjuntos.

export const anotar = (clave, valor) => Sync.ejecutar({ tipo: 'anotar', clave, valor: valor == null ? '' : typeof valor === 'string' ? valor : JSON.stringify(valor) });
export function nota(clave) {
  const v = datos().config[clave]?.valor;
  if (!v) return null;
  try { return JSON.parse(v); } catch { return null; }
}

/** Envío al banco de un pago programado: { por, fecha, grupo } o null. */
export const envioDe = (progId) => nota('envio:' + progId);
export function marcarEnviado(progIds, consecutivo) {
  const grupo = Sync.nuevoId();
  const v = { por: usuarioActual()?.nombre || '', fecha: new Date().toISOString(), grupo, consecutivo: consecutivo || null };
  for (const id of progIds) { const r = anotar('envio:' + id, v); if (r.error) return r; }
  return { ok: true };
}
export function volverPendiente(progIds) {
  for (const id of progIds) { const r = anotar('envio:' + id, ''); if (r.error) return r; }
  return { ok: true };
}

/** Datos del BCR por sociedad: { cedula (la «Cédula jurídica» de la plantilla), ultimo (último consecutivo usado) }. */
export const bancoSoc = (soc) => nota('banco:soc:' + soc) || {};
export const guardarBancoSoc = (soc, cambios) => anotar('banco:soc:' + soc, { ...bancoSoc(soc), ...cambios });
/** Consecutivo que sigue (1-999; después de 999 vuelve a 1). */
export const siguienteConsecutivo = (soc) => ((Number(bancoSoc(soc).ultimo) || 0) % 999) + 1;

/** Datos de un proveedor dentro de un pago aplicado: { documento, archivo: {id, nombre} }. */
export const docLinea = (loteId, codigo) => nota(`doc:${loteId}:${codigo}`) || {};
export const guardarDocLinea = (loteId, codigo, cambios) => anotar(`doc:${loteId}:${codigo}`, { ...docLinea(loteId, codigo), ...cambios });
/** Comprobante general del débito (solo uso interno): { id, nombre }. */
export const adjLote = (loteId) => nota('adj:lote:' + loteId);
export const guardarAdjLote = (loteId, archivo) => anotar('adj:lote:' + loteId, archivo);

/** Comprobante para ver desde un movimiento: el individual del proveedor o, si no hay, el general (interno). */
export function comprobanteDe(m) {
  if (!m?.loteId) return null;
  const f = factura(m.facturaId);
  const ind = f ? docLinea(m.loteId, f.codigo).archivo : null;
  if (ind?.id) return { ...ind, individual: true };
  const gen = adjLote(m.loteId);
  return gen?.id ? { ...gen, individual: false } : null;
}

/** Parte de un lote: solo los proveedores indicados. */
export function subLote(l, codigos) {
  const lineas = l.lineas.filter((ln) => codigos.includes(ln.codigo));
  return { ...l, lineas, total: round2(lineas.reduce((t, x) => t + x.total, 0)) };
}

// ---------------------------------------------------------------- Pago autorizado por Norman
// Norman indica de palabra que se pague o abone una factura y se paga de inmediato.
// Por dentro es un pago programado y aplicado en el mismo momento: queda con comprobante,
// lote y bitácora como cualquier otro, y la nota de la programación dice "Autorizado por Norman".

export const NOTA_NORMAN = 'Autorizado por Norman';
export const NOTA_PAGAR_AHORA = 'Pagar ahora';
export const esAutorizadoNorman = (p) => String(p?.nota || '').startsWith(NOTA_NORMAN) || String(p?.nota || '').startsWith(NOTA_PAGAR_AHORA);

/** Lote de una sola factura (para el .txt del banco y para aplicar). */
export function loteDirecto(f, monto, fecha, descripcion) {
  const lote = { clave: 'directo|' + f.id, sociedad: f.sociedad, moneda: f.moneda, fecha };
  const ln = completarLinea(lote, { codigo: f.codigo, items: [{ p: { id: null, monto: round2(monto) }, f }] });
  if (descripcion != null && String(descripcion).trim()) ln.descripcion = String(descripcion).trim();
  lote.lineas = [ln];
  lote.total = ln.total;
  lote.cuentaDebito = datos().sociedades[f.sociedad]?.[f.moneda === 'USD' ? 'cuentaUSD' : 'cuentaCRC'] || '';
  return lote;
}

export function pagarAutorizadoNorman(f, { monto, fecha, comprobante, nota, descripcion }) {
  return pagoDirecto(f, { monto, fecha, comprobante, descripcion, notaProg: NOTA_PAGAR_AHORA + ' · manual' + (nota ? ' · ' + String(nota).trim() : '') });
}

/** Pagar ahora por banco: se programa para hoy, queda «enviado al banco» y se devuelve el lote para el .txt. */
export function pagarAhoraBanco(f, { monto, fecha, nota, descripcion, consecutivo }) {
  monto = round2(monto);
  const ya = programacionDe(f);
  if (ya) return { error: `Esta factura ya está programada por ${ya.por}. Aplíquela o devuélvala en «Por aplicar» primero.` };
  if (!(monto > 0)) return { error: 'Monto inválido' };
  if (monto > saldo(f) + 0.004) return { error: 'El monto es mayor que el saldo' };
  const progId = Sync.nuevoId();
  const r = Sync.ejecutar({ tipo: 'programar', fechaPago: fecha, nota: NOTA_PAGAR_AHORA + ' · banco' + (nota ? ' · ' + String(nota).trim() : ''), items: [{ progId, facturaId: f.id, monto }] });
  if (r.error) return r;
  const m = marcarEnviado([progId], consecutivo);
  if (m.error) return m;
  if (descripcion != null && String(descripcion).trim()) guardarBorrador([f.sociedad, f.moneda, fecha].join('|'), f.codigo, { descripcion: String(descripcion).trim() });
  return { ok: true, lote: loteDirecto(f, monto, fecha, descripcion) };
}

/** Programa y aplica en un solo paso el pago (o abono) de una factura. */
function pagoDirecto(f, { monto, fecha, comprobante, descripcion, notaProg }) {
  monto = round2(monto);
  const ya = programacionDe(f);
  if (ya) return { error: `Esta factura ya está programada por ${ya.por}. Aplíquela o devuélvala en «Por aplicar» primero.` };
  if (!(monto > 0)) return { error: 'Monto inválido' };
  if (monto > saldo(f) + 0.004) return { error: 'El monto es mayor que el saldo' };
  if (!fecha) return { error: 'Falta la fecha de pago' };
  if (!String(comprobante || '').trim()) return { error: 'Falta el número de comprobante' };
  const progId = Sync.nuevoId();
  const r1 = Sync.ejecutar({
    tipo: 'programar', fechaPago: fecha, nota: notaProg,
    items: [{ progId, facturaId: f.id, monto }],
  });
  if (r1.error) return r1;
  const lote = loteDirecto(f, monto, fecha, descripcion);
  const ln = lote.lineas[0];
  const loteId = Sync.nuevoId();
  const r2 = Sync.ejecutar({
    tipo: 'aplicarLote', loteId, comprobante: String(comprobante).trim(), fecha,
    sociedad: f.sociedad, moneda: f.moneda, progIds: [progId], extras: [],
    lineas: [{ codigo: ln.codigo, total: ln.total, extra: 0, descripcion: ln.descripcion, documentos: [f.documento] }],
  });
  if (r2.error) { Sync.ejecutar({ tipo: 'quitarProgramacion', progId, motivo: 'No se pudo aplicar' }); return r2; }
  return { ok: true, tipo: saldo(f) <= 0.004 ? 'pago' : 'abono', avisos: r2.avisos || [], loteId, codigo: f.codigo };
}

/** Pagos autorizados por Norman de los últimos `dias` días, del más reciente al más antiguo. */
export function autorizadosNorman(dias = 30) {
  const desde = new Date(Date.now() - dias * 864e5).toISOString();
  return Object.values(datos().programaciones)
    .filter((p) => esAutorizadoNorman(p) && p.estado === 'aplicado' && String(p.creado) >= desde)
    .map((p) => ({ p, f: factura(p.facturaId) }))
    .filter((x) => x.f)
    .sort((a, b) => String(b.p.creado).localeCompare(String(a.p.creado)));
}

/** Revisa que el lote se pueda pasar al banco. Devuelve la lista de problemas. */
export function problemasLote(lote) {
  const out = [];
  const mon = lote.moneda === 'USD' ? 'dólares' : 'colones';
  if (!bancoSoc(lote.sociedad).cedula) out.push(`Falta la cédula jurídica de ${nombreSociedad(lote.sociedad)} para el BCR (Configuración → Sociedades).`);
  if (!cfg('controlEncabezado')) out.push('Falta la cédula del empleado que sube el archivo (Configuración → Banco).');
  if (!lote.cuentaDebito) out.push(`Falta la cuenta IBAN en ${mon} de ${nombreSociedad(lote.sociedad)} (Configuración).`);
  else if (L.errorIBAN(lote.cuentaDebito)) out.push(`La cuenta de ${nombreSociedad(lote.sociedad)} no es un IBAN válido.`);
  for (const ln of lote.lineas) {
    if (!ln.iban) out.push(`${ln.prov.nombre} no tiene IBAN en ${mon}.`);
    else if (L.errorIBAN(ln.iban)) out.push(`El IBAN de ${ln.prov.nombre} no es válido.`);
  }
  return out;
}

// ---------------------------------------------------------------- Facturas y catálogos

export const nuevaFactura = (campos) => Sync.ejecutar({ tipo: 'nuevaFactura', factura: { id: Sync.nuevoId(), ...campos } });
/**
 * Abono o nota de crédito registrado desde la ventana de la factura.
 * El script publicado no acepta la operación "movimiento" (su tipo se confunde con el del abono),
 * así que se registra como un pago de una sola factura programado y aplicado a la vez.
 */
export const NOTA_NC = 'Nota de crédito';
export const NOTA_MANUAL = 'Registrado en la factura';
export function registrarMovimiento(facturaId, m) {
  const f = factura(facturaId);
  if (!f) return { error: 'La factura no existe' };
  const nc = m.tipo === 'nc';
  // Nota de crédito: movimiento propio de la factura (no pasa por el banco ni por Lotes).
  if (nc) {
    if (!String(m.comprobante || '').trim()) return { error: 'Falta el número de la nota de crédito' };
    return Sync.ejecutar({ tipo: 'movimiento', tipoMov: 'nc', movId: Sync.nuevoId(), facturaId, monto: round2(m.monto), fecha: m.fecha, comprobante: String(m.comprobante).trim() });
  }
  return pagoDirecto(f, {
    monto: m.monto, fecha: m.fecha, comprobante: m.comprobante, notaProg: nc ? NOTA_NC : NOTA_MANUAL,
    descripcion: nc ? `Nota de credito ${m.comprobante || ''}` : undefined,
  });
}
export const actualizarFactura = (id, cambios) => Sync.ejecutar({ tipo: 'editarFactura', id, cambios });
export const guardarProveedor = (codigo, datosProv, nuevo) => Sync.ejecutar({ tipo: 'guardarProveedor', codigo, datos: datosProv, nuevo });
export const guardarSociedad = (clave, d) => Sync.ejecutar({ tipo: 'guardarSociedad', clave, datos: d });
export const guardarConfig = (cambios) => Sync.ejecutar({ tipo: 'guardarConfig', cambios });

export function siguienteCodigoProveedor() {
  return String(Math.max(0, ...Object.keys(datos().proveedores).map(Number).filter(Number.isFinite)) + 1);
}

/** Convierte el resultado del importador al formato de la hoja. */
export function paqueteImportacion(res) {
  const facturas = [], movimientos = [];
  for (const f of res.facturas) {
    const { movimientos: movs, proveedorExcel, ...resto } = f;
    facturas.push({ ...resto, proveedor: proveedorExcel, creadaPor: 'Excel' });
    for (const m of movs) movimientos.push({ ...m, facturaId: f.id, programadoPor: '', loteId: '' });
  }
  return {
    facturas, movimientos,
    proveedores: Object.values(res.proveedores),
    sociedades: Object.values(res.sociedades).map((s) => ({ clave: s.nombre, corto: s.corto, cuentaCRC: '', cuentaUSD: '', descripcionBanco: '' })),
  };
}
