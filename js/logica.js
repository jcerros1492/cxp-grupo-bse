/*
 * Reglas de negocio de Cuentas por Pagar.
 *
 * Este mismo archivo se usa en dos lugares:
 *   - en el navegador (se carga con <script>, define la variable global Logica);
 *   - en Google Apps Script (se copia como Logica.gs), donde valida cada cambio
 *     antes de guardarlo en la hoja.
 * Por eso no usa import/export ni funciones exclusivas del navegador.
 *
 * Modelo de datos (todas las tablas son objetos indexados por su clave):
 *   facturas[id]          monto original; pagado, saldo y estado se recalculan
 *   movimientos[id]       pagos, abonos, notas de crédito (nc) y cargos (intereses)
 *   programaciones[id]    lo que Norman marca para pagar
 *   proveedores[codigo], sociedades[clave], lotes[id], config[clave] = {clave, valor}
 */
var Logica = (function () {
  'use strict';

  var CLAVES = {
    facturas: 'id', movimientos: 'id', programaciones: 'id', proveedores: 'codigo',
    sociedades: 'clave', lotes: 'id', config: 'clave',
  };
  var TABLAS = Object.keys(CLAVES);

  function vacio() {
    var d = {};
    TABLAS.forEach(function (t) { d[t] = {}; });
    return d;
  }

  function round2(n) { return Math.round((Number(n) + Number.EPSILON) * 100) / 100; }

  // ---------------------------------------------------------------- Índices

  /** Índices internos (no se guardan): movimientos y programación pendiente por factura. */
  function prepararIndices(d) {
    var movs = {}, prog = {};
    Object.keys(d.movimientos).forEach(function (id) {
      var m = d.movimientos[id];
      (movs[m.facturaId] = movs[m.facturaId] || []).push(m);
    });
    Object.keys(movs).forEach(function (k) { movs[k].sort(ordenMov); });
    Object.keys(d.programaciones).forEach(function (id) {
      var p = d.programaciones[id];
      if (p.estado === 'programado') prog[p.facturaId] = p;
    });
    Object.defineProperty(d, '_movs', { value: movs, enumerable: false, configurable: true, writable: true });
    Object.defineProperty(d, '_prog', { value: prog, enumerable: false, configurable: true, writable: true });
    return d;
  }

  function ordenMov(a, b) {
    return String(a.fecha || '').localeCompare(String(b.fecha || '')) || String(a.id).localeCompare(String(b.id));
  }

  function movimientosDe(d, facturaId) { return (d._movs && d._movs[facturaId]) || []; }
  function programacionDe(d, facturaId) { return (d._prog && d._prog[facturaId]) || null; }

  function pagado(d, f) {
    return round2(movimientosDe(d, f.id).reduce(function (s, m) { return s + (m.tipo === 'cargo' ? 0 : Number(m.monto) || 0); }, 0));
  }
  function saldo(d, f) { return round2(Number(f.monto) - pagado(d, f)); }

  function estado(d, f) {
    if (saldo(d, f) <= 0.004) return 'Cancelada';
    if (programacionDe(d, f.id)) return 'Programada';
    if (f.revision) return 'Revisión';
    if (movimientosDe(d, f.id).some(function (m) { return m.tipo !== 'cargo'; })) return 'Abonada';
    return 'Pendiente';
  }

  /** Actualiza las columnas calculadas (pagado, saldo, estado) que se ven en la hoja. */
  function recalcular(d, f) {
    f.pagado = pagado(d, f);
    f.saldo = round2(Number(f.monto) - f.pagado);
    f.estado = estado(d, f);
  }

  function recalcularTodo(d) {
    prepararIndices(d);
    Object.keys(d.facturas).forEach(function (id) { recalcular(d, d.facturas[id]); });
  }

  // ---------------------------------------------------------------- IBAN

  function limpiarIBAN(t) { return String(t == null ? '' : t).replace(/\s+/g, '').toUpperCase(); }

  function errorIBAN(t) {
    var iban = limpiarIBAN(t);
    if (!iban) return 'Vacío';
    if (!/^CR\d{20}$/.test(iban)) return 'Debe ser CR seguido de 20 dígitos (22 caracteres)';
    var num = (iban.slice(4) + iban.slice(0, 4)).replace(/[A-Z]/g, function (c) { return String(c.charCodeAt(0) - 55); });
    var r = 0;
    for (var i = 0; i < num.length; i++) r = (r * 10 + Number(num[i])) % 97;
    return r === 1 ? null : 'Los dígitos de control no coinciden: revise el número';
  }

  // ---------------------------------------------------------------- Operaciones

  /**
   * Aplica una operación sobre los datos. Primero valida y solo después modifica,
   * así una operación inválida no deja cambios a medias.
   * @param d    datos (con índices)
   * @param op   { tipo, ... }
   * @param ctx  { usuario: {nombre, rol}, ahora: número }
   * @returns { ok, error?, avisos?, cambios: [[tabla, clave], ...] }
   */
  function aplicarOp(d, op, ctx) {
    var fn = OPERACIONES[op && op.tipo];
    if (!fn) return { ok: false, error: 'Operación desconocida: ' + (op && op.tipo), cambios: [] };
    var cambios = [];
    var marcar = function (tabla, clave) {
      var e = d[tabla][clave];
      if (e) e.v = ctx.ahora;
      cambios.push([tabla, clave]);
    };
    var r;
    try {
      r = fn(d, op, ctx, marcar) || {};
    } catch (err) {
      return { ok: false, error: String(err && err.message || err), cambios: [] };
    }
    r.cambios = cambios;
    if (r.ok === undefined) r.ok = !r.error;
    return r;
  }

  function quien(ctx) { return (ctx.usuario && ctx.usuario.nombre) || 'Sin usuario'; }

  function nuevoMovimiento(d, m, marcar) {
    d.movimientos[m.id] = m;
    (d._movs[m.facturaId] = d._movs[m.facturaId] || []).push(m);
    d._movs[m.facturaId].sort(ordenMov);
    marcar('movimientos', m.id);
  }

  function tocarFactura(d, f, marcar) {
    recalcular(d, f);
    marcar('facturas', f.id);
  }

  var OPERACIONES = {
    /** Norman: { items: [{progId, facturaId, monto}], fechaPago, nota } */
    programar: function (d, op, ctx, marcar) {
      if (!op.fechaPago) return { error: 'Falta la fecha de pago' };
      if (!op.items || !op.items.length) return { error: 'No se marcó ninguna factura' };
      var avisos = [], validos = [];
      op.items.forEach(function (it) {
        var f = d.facturas[it.facturaId];
        var monto = round2(it.monto);
        if (!f) { avisos.push('Una factura ya no existe'); return; }
        var ya = programacionDe(d, f.id);
        if (d.programaciones[it.progId]) return; // repetida
        if (ya) { avisos.push('Factura ' + f.documento + ': ya estaba programada por ' + ya.por); return; }
        if (!(monto > 0)) { avisos.push('Factura ' + f.documento + ': monto inválido'); return; }
        if (monto > saldo(d, f) + 0.004) { avisos.push('Factura ' + f.documento + ': el monto es mayor que el saldo'); return; }
        validos.push({ it: it, f: f, monto: monto });
      });
      if (!validos.length) return { error: avisos.join(' · ') || 'Nada que programar' };
      validos.forEach(function (x) {
        var p = {
          id: x.it.progId, facturaId: x.f.id, monto: x.monto, fechaPago: op.fechaPago, nota: op.nota || '',
          estado: 'programado', por: quien(ctx), creado: new Date(ctx.ahora).toISOString(),
          cerradoPor: '', cerrado: '', comprobante: '', motivo: '',
        };
        d.programaciones[p.id] = p;
        d._prog[x.f.id] = p;
        marcar('programaciones', p.id);
        tocarFactura(d, x.f, marcar);
      });
      return { ok: true, avisos: avisos };
    },

    /** { progId, motivo } */
    quitarProgramacion: function (d, op, ctx, marcar) {
      var p = d.programaciones[op.progId];
      if (!p) return { error: 'La programación no existe' };
      if (p.estado !== 'programado') return { error: 'Ese pago ya fue ' + (p.estado === 'aplicado' ? 'aplicado' : 'quitado') };
      p.estado = 'devuelto';
      p.cerradoPor = quien(ctx);
      p.cerrado = new Date(ctx.ahora).toISOString();
      p.motivo = op.motivo || 'Quitado';
      delete d._prog[p.facturaId];
      marcar('programaciones', p.id);
      var f = d.facturas[p.facturaId];
      if (f) tocarFactura(d, f, marcar);
      return { ok: true };
    },

    /** Jonathan / Daniel: { loteId, progIds, comprobante, fecha, sociedad, moneda, extras: [{facturaId, monto, codigo}], lineas } */
    aplicarLote: function (d, op, ctx, marcar) {
      if (!op.comprobante) return { error: 'Falta el número de comprobante' };
      if (!op.fecha) return { error: 'Falta la fecha de pago' };
      var avisos = [], aplicar = [];
      (op.progIds || []).forEach(function (pid) {
        var p = d.programaciones[pid];
        if (!p) { avisos.push('Un pago programado ya no existe'); return; }
        var f = d.facturas[p.facturaId];
        if (p.estado !== 'programado') {
          avisos.push('Factura ' + (f ? f.documento : '') + ': ya había sido ' + (p.estado === 'aplicado' ? 'aplicada por ' + p.cerradoPor : 'quitada'));
          return;
        }
        aplicar.push({ p: p, f: f });
      });
      if (!aplicar.length) return { error: avisos.join(' · ') || 'No hay pagos para aplicar' };
      var total = 0, cerrado = new Date(ctx.ahora).toISOString();
      aplicar.forEach(function (x) {
        var s = saldo(d, x.f);
        var monto = round2(Math.min(x.p.monto, s));
        if (monto < x.p.monto - 0.004) avisos.push('Factura ' + x.f.documento + ': se aplicó solo el saldo, ' + monto);
        if (monto > 0) {
          nuevoMovimiento(d, {
            id: op.loteId + '-' + x.p.id, facturaId: x.f.id, tipo: monto >= s - 0.004 ? 'pago' : 'abono', monto: monto,
            fecha: op.fecha, comprobante: op.comprobante, por: quien(ctx), nota: '', programadoPor: x.p.por, loteId: op.loteId,
          }, marcar);
          total += monto;
        }
        x.p.estado = 'aplicado';
        x.p.cerradoPor = quien(ctx);
        x.p.cerrado = cerrado;
        x.p.comprobante = op.comprobante;
        delete d._prog[x.f.id];
        marcar('programaciones', x.p.id);
        tocarFactura(d, x.f, marcar);
      });
      (op.extras || []).forEach(function (e) {
        var f = d.facturas[e.facturaId];
        var monto = round2(e.monto);
        if (!f || !(monto > 0)) return;
        nuevoMovimiento(d, {
          id: op.loteId + '-x-' + f.id, facturaId: f.id, tipo: 'cargo', monto: monto, fecha: op.fecha,
          comprobante: op.comprobante, por: quien(ctx), nota: 'Intereses u otros cargos', programadoPor: '', loteId: op.loteId,
        }, marcar);
        total += monto;
        tocarFactura(d, f, marcar);
      });
      d.lotes[op.loteId] = {
        id: op.loteId, sociedad: op.sociedad || '', moneda: op.moneda || '', fecha: op.fecha, comprobante: op.comprobante,
        total: round2(total), aplicadoPor: quien(ctx), aplicado: cerrado, detalle: JSON.stringify(op.lineas || []),
      };
      marcar('lotes', op.loteId);
      return { ok: true, avisos: avisos };
    },

    /** Abono, pago o nota de crédito manual: { movId, facturaId, tipo, monto, fecha, comprobante, nota } */
    movimiento: function (d, op, ctx, marcar) {
      var f = d.facturas[op.facturaId];
      var monto = round2(op.monto);
      if (!f) return { error: 'La factura no existe' };
      if (d.movimientos[op.movId]) return { ok: true };
      if (['abono', 'pago', 'nc'].indexOf(op.tipo) < 0) return { error: 'Tipo de movimiento inválido' };
      if (!(monto > 0)) return { error: 'Monto inválido' };
      if (monto > saldo(d, f) + 0.004) return { error: 'El monto es mayor que el saldo de la factura' };
      if (!op.fecha) return { error: 'Falta la fecha' };
      nuevoMovimiento(d, {
        id: op.movId, facturaId: f.id, tipo: op.tipo, monto: monto, fecha: op.fecha, comprobante: op.comprobante || '',
        por: quien(ctx), nota: op.nota || '', programadoPor: '', loteId: '',
      }, marcar);
      tocarFactura(d, f, marcar);
      return { ok: true };
    },

    /** { factura: {id, sociedad, codigo, documento, moneda, fechaFactura, fechaVence, monto, observaciones} } */
    nuevaFactura: function (d, op, ctx, marcar) {
      var n = op.factura || {};
      if (d.facturas[n.id]) return { ok: true };
      if (!n.sociedad || !n.codigo || !n.documento) return { error: 'Faltan datos de la factura' };
      if (!(round2(n.monto) > 0)) return { error: 'Monto inválido' };
      var dup = Object.keys(d.facturas).some(function (id) {
        var f = d.facturas[id];
        return f.sociedad === n.sociedad && String(f.codigo) === String(n.codigo) && String(f.documento) === String(n.documento) && f.moneda === n.moneda;
      });
      if (dup) return { error: 'Ya existe la factura ' + n.documento + ' de ese proveedor en esa sociedad' };
      var prov = d.proveedores[n.codigo];
      var f = {
        id: n.id, sociedad: n.sociedad, codigo: String(n.codigo), proveedor: prov ? prov.nombre : '', documento: String(n.documento),
        fechaFactura: n.fechaFactura || '', fechaVence: n.fechaVence || '', venceEstimado: false, moneda: n.moneda === 'USD' ? 'USD' : 'CRC',
        monto: round2(n.monto), revision: false, observaciones: n.observaciones || '', creadaPor: quien(ctx),
      };
      d.facturas[f.id] = f;
      tocarFactura(d, f, marcar);
      return { ok: true };
    },

    /** { id, cambios: {fechaFactura, fechaVence, observaciones, revision} } */
    editarFactura: function (d, op, ctx, marcar) {
      var f = d.facturas[op.id];
      if (!f) return { error: 'La factura no existe' };
      var c = op.cambios || {};
      ['fechaFactura', 'observaciones'].forEach(function (k) { if (k in c) f[k] = c[k] || ''; });
      if ('revision' in c) f.revision = !!c.revision;
      if ('fechaVence' in c && (c.fechaVence || '') !== (f.fechaVence || '')) { f.fechaVence = c.fechaVence || ''; f.venceEstimado = false; }
      tocarFactura(d, f, marcar);
      return { ok: true };
    },

    /** { codigo, nuevo, datos: {nombre, alias, ibanCRC, ibanUSD, plazo} } */
    guardarProveedor: function (d, op, ctx, marcar) {
      var codigo = String(op.codigo || '').trim();
      var x = op.datos || {};
      if (!codigo) return { error: 'Falta el código' };
      if (op.nuevo && d.proveedores[codigo]) return { error: 'Ya existe un proveedor con el código ' + codigo };
      if (!String(x.nombre || '').trim()) return { error: 'Falta el nombre' };
      var ib = { ibanCRC: limpiarIBAN(x.ibanCRC), ibanUSD: limpiarIBAN(x.ibanUSD) };
      if (ib.ibanCRC && errorIBAN(ib.ibanCRC)) return { error: 'El IBAN en colones no es válido' };
      if (ib.ibanUSD && errorIBAN(ib.ibanUSD)) return { error: 'El IBAN en dólares no es válido' };
      d.proveedores[codigo] = {
        codigo: codigo, nombre: String(x.nombre).trim(), alias: String(x.alias || '').trim(),
        ibanCRC: ib.ibanCRC, ibanUSD: ib.ibanUSD, plazo: Number(x.plazo) || 30,
      };
      marcar('proveedores', codigo);
      return { ok: true };
    },

    /** { clave, datos: {corto, cuentaCRC, cuentaUSD, descripcionBanco} } */
    guardarSociedad: function (d, op, ctx, marcar) {
      var s = d.sociedades[op.clave];
      if (!s) return { error: 'La sociedad no existe' };
      var x = op.datos || {};
      var c = { cuentaCRC: limpiarIBAN(x.cuentaCRC), cuentaUSD: limpiarIBAN(x.cuentaUSD) };
      if ((c.cuentaCRC && errorIBAN(c.cuentaCRC)) || (c.cuentaUSD && errorIBAN(c.cuentaUSD))) return { error: 'Hay un IBAN inválido' };
      s.corto = String(x.corto || s.corto || '').trim();
      s.descripcionBanco = String(x.descripcionBanco || '').trim();
      s.cuentaCRC = c.cuentaCRC;
      s.cuentaUSD = c.cuentaUSD;
      marcar('sociedades', s.clave);
      return { ok: true };
    },

    /** { cambios: {clienteBanco, controlEncabezado} } */
    guardarConfig: function (d, op, ctx, marcar) {
      var c = op.cambios || {};
      if (c.clienteBanco && !/^\d{1,15}$/.test(c.clienteBanco)) return { error: 'El código de cliente debe tener solo dígitos' };
      if (c.controlEncabezado && !/^\d{1,12}$/.test(c.controlEncabezado)) return { error: 'El campo de control debe tener hasta 12 dígitos' };
      ['clienteBanco', 'controlEncabezado'].forEach(function (k) {
        if (k in c) { d.config[k] = { clave: k, valor: String(c[k] || '') }; marcar('config', k); }
      });
      return { ok: true };
    },
  };

  function cfg(d, clave) { var c = d.config[clave]; return c ? String(c.valor == null ? '' : c.valor) : ''; }

  return {
    CLAVES: CLAVES, TABLAS: TABLAS, vacio: vacio, round2: round2,
    prepararIndices: prepararIndices, recalcular: recalcular, recalcularTodo: recalcularTodo,
    movimientosDe: movimientosDe, programacionDe: programacionDe, pagado: pagado, saldo: saldo, estado: estado,
    limpiarIBAN: limpiarIBAN, errorIBAN: errorIBAN, aplicarOp: aplicarOp, cfg: cfg,
  };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = Logica;
