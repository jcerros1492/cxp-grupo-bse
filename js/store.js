// Estado de la aplicación y reglas de negocio.
//
// Versión 1: los datos se guardan en el navegador (localStorage) de cada computadora.
// Toda lectura/escritura pasa por este módulo para poder cambiarlo luego por una
// base de datos compartida sin tocar las pantallas.

import { round2, uid, hoyISO, diasDesde, errorIBAN, limpiarIBAN } from './util.js';

const CLAVE = 'cxp-datos-v1';
const CLAVE_USUARIO = 'cxp-usuario';

export const USUARIOS = [
  { id: 'norman', nombre: 'Norman Alpízar', rol: 'Aprueba pagos' },
  { id: 'jonathan', nombre: 'Jonathan Cerros', rol: 'Registra y aplica' },
  { id: 'daniel', nombre: 'Daniel Barquero', rol: 'Aplica pagos' },
];

function vacio() {
  return {
    version: 1,
    importado: null,
    facturas: [],
    proveedores: {},
    sociedades: {},
    programaciones: [],
    lotes: [],
    borradores: {}, // descripciones y cargos extra editados en "Por aplicar"
    config: { clienteBanco: '', controlEncabezado: '' },
  };
}

export let datos = cargar();

function cargar() {
  try {
    const t = localStorage.getItem(CLAVE);
    if (!t) return vacio();
    return { ...vacio(), ...JSON.parse(t) };
  } catch {
    return vacio();
  }
}

const oyentes = new Set();
export const suscribir = (fn) => oyentes.add(fn);

export function guardar() {
  try {
    localStorage.setItem(CLAVE, JSON.stringify(datos));
  } catch (e) {
    alert('No se pudo guardar en este navegador: ' + e.message);
  }
  oyentes.forEach((fn) => fn());
}

export function reemplazarDatos(nuevos) {
  datos = { ...vacio(), ...nuevos };
  guardar();
}

// ---------- Usuario actual ----------

export function usuarioActual() {
  const id = localStorage.getItem(CLAVE_USUARIO);
  return USUARIOS.find((u) => u.id === id) || null;
}
export function cambiarUsuario(id) {
  localStorage.setItem(CLAVE_USUARIO, id);
  oyentes.forEach((fn) => fn());
}
const quien = () => usuarioActual()?.nombre || 'Sin usuario';

// ---------- Importación ----------

/** Carga el resultado del Excel. Conserva IBAN, alias y cuentas ya configurados. */
export function cargarImportacion(res) {
  const provAnt = datos.proveedores, socAnt = datos.sociedades;
  for (const [cod, p] of Object.entries(res.proveedores)) {
    const a = provAnt[cod];
    if (a) Object.assign(p, { alias: a.alias, ibanCRC: a.ibanCRC, ibanUSD: a.ibanUSD, plazo: a.plazo ?? 30 });
  }
  for (const [k, s] of Object.entries(res.sociedades)) {
    const a = socAnt[k];
    if (a) Object.assign(s, { corto: a.corto, cuentaCRC: a.cuentaCRC, cuentaUSD: a.cuentaUSD, descripcionBanco: a.descripcionBanco });
  }
  for (const [cod, p] of Object.entries(provAnt)) if (!res.proveedores[cod]) res.proveedores[cod] = p;
  datos = {
    ...vacio(),
    config: datos.config,
    importado: new Date().toISOString(),
    facturas: res.facturas,
    proveedores: res.proveedores,
    sociedades: res.sociedades,
  };
  guardar();
}

// ---------- Cálculos ----------

export const factura = (id) => datos.facturas.find((f) => f.id === id);
export const proveedor = (codigo) => datos.proveedores[codigo] || { codigo, nombre: codigo, alias: '' };
export const nombreSociedad = (s) => datos.sociedades[s]?.corto || s;

export function pagado(f) {
  return round2(f.movimientos.reduce((s, m) => s + (m.tipo === 'cargo' ? 0 : m.monto), 0));
}
export const saldo = (f) => round2(f.monto - pagado(f));

/** Programación pendiente (aún no aplicada) de una factura */
export const programacionDe = (f) =>
  datos.programaciones.find((p) => p.facturaId === f.id && p.estado === 'programado');

export function estado(f) {
  const s = saldo(f);
  if (s <= 0.004) return 'Cancelada';
  if (programacionDe(f)) return 'Programada';
  if (f.revision) return 'Revisión';
  if (f.movimientos.some((m) => m.tipo !== 'cargo')) return 'Abonada';
  return 'Pendiente';
}

export const diasVencida = (f) => (f.fechaVence ? diasDesde(f.fechaVence) : null);

export const conSaldo = () => datos.facturas.filter((f) => saldo(f) > 0.004);

export function totales(lista) {
  const t = { CRC: 0, USD: 0, n: lista.length };
  for (const f of lista) t[f.moneda] = round2(t[f.moneda] + saldo(f));
  return t;
}

// ---------- Programar pagos (Norman) ----------

/** items: [{facturaId, monto}] */
export function programarPagos(items, fechaPago, nota = '') {
  const errores = [];
  for (const it of items) {
    const f = factura(it.facturaId);
    const s = saldo(f);
    if (!(it.monto > 0)) errores.push(`Factura ${f.documento}: monto inválido`);
    else if (it.monto > s + 0.004) errores.push(`Factura ${f.documento}: el monto es mayor que el saldo`);
    else if (programacionDe(f)) errores.push(`Factura ${f.documento}: ya está programada`);
  }
  if (!fechaPago) errores.push('Falta la fecha de pago');
  if (errores.length) return errores;
  const ahora = new Date().toISOString();
  for (const it of items) {
    datos.programaciones.push({
      id: uid(), facturaId: it.facturaId, monto: round2(it.monto), fechaPago, nota,
      estado: 'programado', por: quien(), creado: ahora,
    });
  }
  guardar();
  return [];
}

export function quitarProgramacion(id, motivo = 'Quitada') {
  const p = datos.programaciones.find((x) => x.id === id);
  if (!p) return;
  p.estado = 'devuelto';
  p.cerradoPor = quien();
  p.cerrado = new Date().toISOString();
  p.motivo = motivo;
  guardar();
}

export function cambiarMontoProgramado(id, monto) {
  const p = datos.programaciones.find((x) => x.id === id);
  const f = factura(p.facturaId);
  if (!(monto > 0) || monto > saldo(f) + 0.004) return 'Monto inválido o mayor que el saldo';
  p.monto = round2(monto);
  guardar();
  return null;
}

// ---------- Lotes por aplicar (Jonathan y Daniel) ----------

/** Agrupa lo programado en lotes: sociedad + moneda + fecha; dentro, una línea por proveedor. */
export function lotesPendientes() {
  const mapa = new Map();
  for (const p of datos.programaciones.filter((x) => x.estado === 'programado')) {
    const f = factura(p.facturaId);
    if (!f) continue;
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
    l.cuentaDebito = datos.sociedades[l.sociedad]?.[l.moneda === 'USD' ? 'cuentaUSD' : 'cuentaCRC'] || '';
  }
  return lotes;
}

function completarLinea(lote, ln) {
  const prov = proveedor(ln.codigo);
  const b = datos.borradores[lote.clave + '|' + ln.codigo] || {};
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

export function guardarBorrador(loteClave, codigo, cambios) {
  const k = loteClave + '|' + codigo;
  datos.borradores[k] = { ...(datos.borradores[k] || {}), ...cambios };
  guardar();
}

/** Registra el pago de todo el lote con un comprobante. */
export function aplicarLote(lote, { comprobante, fecha }) {
  if (!comprobante) return 'Falta el número de comprobante';
  if (!fecha) return 'Falta la fecha de pago';
  const ahora = new Date().toISOString();
  for (const ln of lote.lineas) {
    for (const { p, f } of ln.items) {
      const monto = Math.min(p.monto, saldo(f));
      if (monto > 0) {
        f.movimientos.push({
          id: uid(), tipo: monto >= saldo(f) - 0.004 ? 'pago' : 'abono', monto: round2(monto), fecha, comprobante,
          por: quien(), nota: '', programadoPor: p.por,
        });
      }
      Object.assign(p, { estado: 'aplicado', cerradoPor: quien(), cerrado: ahora, comprobante });
    }
    if (ln.extra > 0) {
      const f = ln.items[0].f;
      f.movimientos.push({ id: uid(), tipo: 'cargo', monto: ln.extra, fecha, comprobante, por: quien(), nota: 'Intereses u otros cargos pagados con esta factura' });
    }
    delete datos.borradores[lote.clave + '|' + ln.codigo];
  }
  datos.lotes.push({
    id: uid(), sociedad: lote.sociedad, moneda: lote.moneda, fecha, comprobante, total: lote.total, aplicadoPor: quien(), aplicado: ahora,
    lineas: lote.lineas.map((ln) => ({ codigo: ln.codigo, total: ln.total, extra: ln.extra, descripcion: ln.descripcion, documentos: ln.items.map((x) => x.f.documento) })),
  });
  guardar();
  return null;
}

/** Revisa que el lote se pueda pasar al banco. Devuelve la lista de problemas. */
export function problemasLote(lote) {
  const out = [];
  if (!datos.config.clienteBanco) out.push('Falta el código de cliente del banco (Configuración).');
  if (!lote.cuentaDebito) out.push(`Falta la cuenta IBAN en ${lote.moneda === 'USD' ? 'dólares' : 'colones'} de ${nombreSociedad(lote.sociedad)} (Configuración).`);
  else if (errorIBAN(lote.cuentaDebito)) out.push(`La cuenta de ${nombreSociedad(lote.sociedad)} no es un IBAN válido.`);
  for (const ln of lote.lineas) {
    if (!ln.iban) out.push(`${ln.prov.nombre} no tiene IBAN en ${lote.moneda === 'USD' ? 'dólares' : 'colones'}.`);
    else if (errorIBAN(ln.iban)) out.push(`El IBAN de ${ln.prov.nombre} no es válido.`);
  }
  return out;
}

// ---------- Facturas, notas de crédito y proveedores ----------

export function nuevaFactura(campos) {
  const dup = datos.facturas.find((f) => f.sociedad === campos.sociedad && f.codigo === campos.codigo && f.documento === campos.documento && f.moneda === campos.moneda);
  if (dup) return `Ya existe la factura ${campos.documento} de ese proveedor en esa sociedad`;
  datos.facturas.push({ id: uid(), revision: false, observaciones: '', movimientos: [], venceEstimado: false, ...campos, monto: round2(campos.monto), creadaPor: quien(), creada: new Date().toISOString() });
  guardar();
  return null;
}

export function registrarMovimiento(facturaId, { tipo, monto, fecha, comprobante, nota }) {
  const f = factura(facturaId);
  if (!(monto > 0)) return 'Monto inválido';
  if (monto > saldo(f) + 0.004) return 'El monto es mayor que el saldo de la factura';
  if (!fecha) return 'Falta la fecha';
  f.movimientos.push({ id: uid(), tipo, monto: round2(monto), fecha, comprobante: comprobante || '', nota: nota || '', por: quien() });
  guardar();
  return null;
}

export function actualizarFactura(id, cambios) {
  Object.assign(factura(id), cambios);
  guardar();
}

export function actualizarProveedor(codigo, cambios) {
  for (const k of ['ibanCRC', 'ibanUSD']) if (k in cambios) cambios[k] = limpiarIBAN(cambios[k]);
  datos.proveedores[codigo] = { ...proveedor(codigo), ...cambios };
  guardar();
}

export function siguienteCodigoProveedor() {
  return String(Math.max(0, ...Object.keys(datos.proveedores).map(Number).filter(Number.isFinite)) + 1);
}

export function actualizarSociedad(clave, cambios) {
  for (const k of ['cuentaCRC', 'cuentaUSD']) if (k in cambios) cambios[k] = limpiarIBAN(cambios[k]);
  datos.sociedades[clave] = { ...datos.sociedades[clave], ...cambios };
  guardar();
}

export function actualizarConfig(cambios) {
  datos.config = { ...datos.config, ...cambios };
  guardar();
}

export { hoyISO };
