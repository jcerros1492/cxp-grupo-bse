// Pantallas y eventos de la aplicación.
import * as S from './store.js';
import * as Sync from './sync.js';
import { importarExcel } from './importer.js';
import { construirArchivo } from './banco.js';
import {
  money, fmtNum, parseNum, fechaHora, fmtFecha, hoyISO, esc, errorIBAN, fmtIBAN, sumarDias, round2, MONEDAS,
} from './util.js';

const $app = document.getElementById('app');
const $nav = document.getElementById('nav');
const $user = document.getElementById('usuario');
const $dialog = document.getElementById('dialogo');

// Estado temporal de pantallas (no se guarda)
const ui = {
  prog: { soc: '', mon: '', q: '', vencidas: false, sel: new Map(), fecha: '', verProgramado: true },
  fact: { soc: '', mon: '', est: 'saldo', q: '', limite: 150 },
  prov: { q: '', filtro: '', abierto: null },
};

// ---------------------------------------------------------------- Navegación

const RUTAS = {
  programar: { titulo: 'Programar pagos', vista: vistaProgramar },
  aplicar: { titulo: 'Por aplicar', vista: vistaAplicar },
  facturas: { titulo: 'Facturas', vista: vistaFacturas },
  proveedores: { titulo: 'Proveedores', vista: vistaProveedores },
  tablero: { titulo: 'Tablero', vista: vistaTablero },
  config: { titulo: 'Configuración', vista: vistaConfig },
};

function menuPara(usuario) {
  if (usuario?.rol === 'aprueba') return ['programar', 'tablero', 'facturas'];
  return ['aplicar', 'programar', 'facturas', 'proveedores', 'tablero', 'config'];
}

function rutaActual() {
  const r = location.hash.replace(/^#\/?/, '').split('?')[0];
  if (RUTAS[r]) return r;
  return S.esNorman() ? 'programar' : 'aplicar';
}

const ROLES = { admin: 'Administra y aplica', aplica: 'Aplica pagos', aprueba: 'Aprueba pagos' };

function indicadorRed() {
  const e = Sync.estado, n = e.pendientes.length;
  const pend = n ? ` · ${n} ${n === 1 ? 'cambio por enviar' : 'cambios por enviar'}` : '';
  if (e.red === 'offline') return `<span class="red off" title="Los cambios se guardan aquí y se envían al volver la conexión">Sin internet${pend}</span>`;
  if (e.red === 'sync') return `<span class="red sync">Sincronizando…${pend}</span>`;
  if (e.red === 'error') return `<button class="red err" data-act="sync-ahora" title="${esc(e.error)}">No se pudo sincronizar${pend} · reintentar</button>`;
  return `<button class="red ok" data-act="sync-ahora" title="Sincronizar ahora">Sincronizado${e.ultimoSync ? ' ' + fechaHora(e.ultimoSync) : ''}${pend}</button>`;
}

function render() {
  const u = S.usuarioActual();
  const ruta = rutaActual();
  const foco = guardarFoco();
  $nav.innerHTML = u ? menuPara(u).map((r) => {
    const n = r === 'aplicar' ? S.programacionesPendientes().length : 0;
    return `<a href="#/${r}" class="${r === ruta ? 'activo' : ''}" ${r === ruta ? 'aria-current="page"' : ''}>${RUTAS[r].titulo}${n ? ` <span class="burbuja">${n}</span>` : ''}</a>`;
  }).join('') : '';
  $user.innerHTML = u
    ? `${indicadorRed()}<button class="usuario-info" data-act="cuenta" title="Mi cuenta"><span class="avatar" aria-hidden="true">${esc(u.nombre.split(' ').map((p) => p[0]).join('').slice(0, 2))}</span><span><strong>${esc(u.nombre)}</strong><small>${esc(ROLES[u.rol] || u.rol)}</small></span></button>`
    : '';
  document.title = `${RUTAS[ruta].titulo} · Cuentas por Pagar`;

  if (!Sync.estado.cargado) { $app.innerHTML = '<p class="vacio">Cargando…</p>'; return; }
  if (!Sync.conexion()) { $app.innerHTML = vistaConexion(); return; }
  if (!u) { $app.innerHTML = vistaLogin(); return; }
  if (u.debeCambiar) { $nav.innerHTML = ''; $app.innerHTML = vistaCrearContrasena(); return; }
  if (!S.hayDatos() && ruta !== 'config') { $app.innerHTML = vistaBienvenida(); return; }
  $app.innerHTML = RUTAS[ruta].vista();
  despuesDeRender(ruta);
  restaurarFoco(foco);
}

/** Al llegar cambios de otra persona se redibuja la pantalla: se conserva el campo donde se escribe. */
function guardarFoco() {
  const el = document.activeElement;
  if (!el || !$app.contains(el) || !el.matches('input, select, textarea')) return null;
  let sel = null;
  if (el.id) sel = '#' + CSS.escape(el.id);
  else if (el.dataset.in) sel = `[data-in="${el.dataset.in}"]` + (el.dataset.v ? `[data-v="${CSS.escape(el.dataset.v)}"]` : '');
  else if (el.name && el.form?.dataset.form) sel = `[data-form="${el.form.dataset.form}"] [name="${el.name}"]`;
  return { sel, valor: el.value, ini: el.selectionStart, fin: el.selectionEnd };
}
function restaurarFoco(f) {
  if (!f?.sel) return;
  const el = $app.querySelector(f.sel);
  if (!el) return;
  if (el.type !== 'checkbox' && el.value !== f.valor) el.value = f.valor;
  el.focus();
  try { if (f.ini != null) el.setSelectionRange(f.ini, f.fin); } catch { /* campo sin selección */ }
}

const nLineas = (l) => { try { return JSON.parse(l.detalle || '[]').length; } catch { return ''; } };

window.addEventListener('hashchange', () => { render(); window.scrollTo(0, 0); });
S.suscribir(render);

// ---------------------------------------------------------------- Piezas comunes

const chip = (act, valor, actual, texto) =>
  `<button type="button" class="chip ${valor === actual ? 'on' : ''}" data-act="${act}" data-v="${esc(valor)}" aria-pressed="${valor === actual}">${texto}</button>`;

function chipsSociedad(act, actual) {
  return chip(act, '', actual, 'Todas') + Object.keys(S.sociedades()).map((s) => chip(act, s, actual, esc(S.nombreSociedad(s)))).join('');
}

function etiquetaDias(f) {
  const d = S.diasVencida(f);
  if (d == null) return '';
  const est = f.venceEstimado ? ' <span class="est" title="Vencimiento estimado: fecha de factura + 30 días">EST</span>' : '';
  if (d > 0) return `<span class="dias venc">vencida ${d} d</span>${est}`;
  if (d === 0) return `<span class="dias hoy">vence hoy</span>${est}`;
  return `<span class="dias">en ${-d} d</span>${est}`;
}

function etiquetaEstado(e) {
  const cls = { Pendiente: 'pend', Programada: 'prog', Abonada: 'abon', Cancelada: 'canc', 'Revisión': 'rev' }[e] || '';
  return `<span class="estado ${cls}">${e}</span>`;
}

const totalesTxt = (t) => [t.CRC ? money(t.CRC, 'CRC') : '', t.USD ? money(t.USD, 'USD') : ''].filter(Boolean).join(' · ') || '—';

// ---------------------------------------------------------------- Inicio

function vistaLogin() {
  return `<section class="centro">
    <h1>Cuentas por Pagar</h1>
    <p class="sub">Entre con su usuario y contraseña. Si es la primera vez, use la contraseña temporal que le dio Jonathan.</p>
    <form class="form-login" data-form="login">
      <label>Usuario <input name="usuario" required autocomplete="username" autocapitalize="none" spellcheck="false" placeholder="Ej. norman"></label>
      <label>Contraseña <input name="contrasena" type="password" required autocomplete="current-password"></label>
      <button class="btn primario grande" type="submit">Entrar</button>
    </form>
    <p class="nota">Después de entrar la primera vez, la aplicación funciona también sin internet en esta computadora.</p>
    <p class="nota"><button class="btn-link" data-act="conexion">Configuración de la conexión</button></p>
  </section>`;
}

/** Primera vez en esta computadora: dirección del servicio y su contraseña. */
function vistaConexion() {
  return `<section class="centro">
    <h1>Configurar esta computadora</h1>
    <p class="sub">Para conectar esta computadora con la hoja de Google se necesita la dirección del servicio y la contraseña del servicio. Las tiene Jonathan.</p>
    ${formConexion()}
  </section>`;
}

function formConexion() {
  return `<form class="form-login" data-form="conexion">
      <label>Dirección del servicio <input name="url" type="url" required placeholder="https://script.google.com/macros/s/…/exec" autocomplete="off"></label>
      <label>Contraseña del servicio <input name="clave" type="password" required autocomplete="off"></label>
      <button class="btn primario grande" type="submit">Conectar</button>
    </form>`;
}

function dialogoConexion() {
  const c = Sync.conexion();
  abrirDialogo(`
    <header class="dlg-cab"><div><h2>Conexión con Google Sheets</h2>
      <p class="sub">${c ? 'Esta computadora está conectada. Para cambiar la conexión escriba la nueva dirección y la contraseña del servicio.' : 'Esta computadora no está conectada.'}</p></div>
      <button class="btn-cerrar" data-act="cerrar" aria-label="Cerrar">×</button></header>
    <div class="dlg-cuerpo">${formConexion()}</div>`);
}

function vistaCrearContrasena() {
  const u = S.usuarioActual();
  return `<section class="centro">
    <h1>Hola, ${esc(u.nombre.split(' ')[0])}</h1>
    <p class="sub">Entró con una contraseña temporal. Cree su propia contraseña para continuar; solo usted la va a conocer.</p>
    <form class="form-login" data-form="nueva-clave">
      <label>Nueva contraseña <input name="nueva" type="password" required minlength="8" autocomplete="new-password"><span class="ayuda">Al menos 8 caracteres</span></label>
      <label>Repítala <input name="repetir" type="password" required minlength="8" autocomplete="new-password"></label>
      <button class="btn primario grande" type="submit">Guardar contraseña</button>
    </form>
    <p class="nota"><button class="btn-link" data-act="salir">Salir</button></p>
  </section>`;
}

function dialogoCuenta() {
  const u = S.usuarioActual();
  abrirDialogo(`
    <header class="dlg-cab"><div><h2>${esc(u.nombre)}</h2><p class="sub">Usuario <strong>${esc(u.usuario)}</strong> · ${esc(ROLES[u.rol] || u.rol)}</p></div>
      <button class="btn-cerrar" data-act="cerrar" aria-label="Cerrar">×</button></header>
    <div class="dlg-cuerpo">
      <h3>Cambiar contraseña</h3>
      <form class="form-grid" data-form="cambiar-clave">
        <label>Contraseña actual <input name="actual" type="password" required autocomplete="current-password"></label>
        <label>Nueva contraseña <input name="nueva" type="password" required minlength="8" autocomplete="new-password"></label>
        <label>Repítala <input name="repetir" type="password" required minlength="8" autocomplete="new-password"></label>
        <div class="ancho"><button class="btn primario" type="submit">Cambiar contraseña</button></div>
      </form>
      <h3>Sesión</h3>
      <p class="sub">La sesión queda abierta en esta computadora, también sin internet, hasta que la cierre.</p>
      <button class="btn" data-act="salir">Cerrar sesión</button>
    </div>`);
}

function vistaBienvenida() {
  if (S.esAdmin()) {
    return `<section class="centro">
      <h1>Cargar las cuentas por pagar</h1>
      <p class="sub">La hoja de Google está vacía. Elija el Excel de cuentas por pagar: se convierte aquí y se guarda en Google Sheets. Toma un minuto.</p>
      <label class="btn primario grande">Elegir archivo de Excel<input type="file" accept=".xlsx,.xls" data-in="importar" hidden></label>
      <p class="nota" id="progreso-import"></p>
    </section>`;
  }
  return `<section class="centro">
    <h1>Todavía no hay datos</h1>
    <p class="sub">Jonathan debe cargar el Excel por primera vez. Cuando lo haga, los datos aparecerán aquí solos.</p>
    <button class="btn" data-act="sync-ahora">Revisar ahora</button>
  </section>`;
}

// ---------------------------------------------------------------- Programar pagos (Norman)

function facturasParaProgramar() {
  const { soc, mon, q, vencidas } = ui.prog;
  const texto = q.trim().toLowerCase();
  return S.conSaldo().filter((f) => {
    if (soc && f.sociedad !== soc) return false;
    if (mon && f.moneda !== mon) return false;
    if (vencidas && !(S.diasVencida(f) >= 0)) return false;
    if (texto) {
      const p = S.proveedor(f.codigo);
      if (!(`${p.nombre} ${p.alias} ${f.documento}`.toLowerCase().includes(texto))) return false;
    }
    return true;
  });
}

function vistaProgramar() {
  const P = ui.prog;
  if (!P.fecha) P.fecha = hoyISO();
  const lista = facturasParaProgramar();
  const grupos = new Map();
  for (const f of lista) {
    const k = f.codigo + '|' + f.sociedad + '|' + f.moneda;
    if (!grupos.has(k)) grupos.set(k, []);
    grupos.get(k).push(f);
  }
  const ordenados = [...grupos.values()]
    .map((fs) => fs.sort((a, b) => (a.fechaVence || '').localeCompare(b.fechaVence || '')))
    .sort((a, b) => S.proveedor(a[0].codigo).nombre.localeCompare(S.proveedor(b[0].codigo).nombre));

  const programado = S.programacionesPendientes()
    .map((p) => ({ p, f: S.factura(p.facturaId) })).filter((x) => x.f)
    .sort((a, b) => a.p.fechaPago.localeCompare(b.p.fechaPago));

  return `
  <div class="titulo-vista">
    <div>
      <h1>Programar pagos</h1>
      <p class="sub">Marque las facturas, escriba cuánto pagar de cada una (si es menos del saldo, queda como abono) y elija la fecha.</p>
    </div>
  </div>

  <div class="filtros">
    <div class="fila-chips" role="group" aria-label="Sociedad">${chipsSociedad('p-soc', P.soc)}</div>
    <div class="fila-chips" role="group" aria-label="Moneda">${chip('p-mon', '', P.mon, 'Ambas monedas')}${chip('p-mon', 'CRC', P.mon, 'Colones')}${chip('p-mon', 'USD', P.mon, 'Dólares')}</div>
    <div class="fila-busqueda">
      <label class="buscar"><span class="vh">Buscar</span><input type="search" placeholder="Buscar proveedor o número de factura" value="${esc(P.q)}" data-in="p-q"></label>
      <label class="check"><input type="checkbox" data-in="p-venc" ${P.vencidas ? 'checked' : ''}> Solo vencidas o que vencen hoy</label>
    </div>
  </div>

  <p class="resumen-lista">${lista.length} facturas con saldo · ${totalesTxt(S.totales(lista))}</p>

  <div class="grupos">
    ${ordenados.length ? ordenados.map(grupoProveedor).join('') : '<p class="vacio">No hay facturas con saldo para estos filtros.</p>'}
  </div>

  ${programado.length ? `
  <section class="bloque">
    <h2>Ya programado <span class="sub-inline">${programado.length} facturas, esperando que Jonathan o Daniel las apliquen</span></h2>
    <div class="tabla-scroll"><table class="tabla">
      <thead><tr><th>Pagar el</th><th>Sociedad</th><th>Proveedor</th><th>Factura</th><th class="num">A pagar</th><th>Programó</th><th></th></tr></thead>
      <tbody>${programado.map(({ p, f }) => `<tr>
        <td>${fmtFecha(p.fechaPago)}</td><td>${esc(S.nombreSociedad(f.sociedad))}</td>
        <td>${esc(S.proveedor(f.codigo).nombre)}</td><td>${esc(f.documento)}${p.monto < S.saldo(f) - 0.004 ? ' <span class="tag-abono">abono</span>' : ''}</td>
        <td class="num">${money(p.monto, f.moneda)}</td><td>${esc(p.por)}</td>
        <td><button class="btn chico" data-act="p-quitar" data-v="${p.id}">Quitar</button></td></tr>`).join('')}</tbody>
    </table></div>
  </section>` : ''}

  <div class="barra-envio" id="barra-envio">
    <div class="barra-interior">
      <div id="resumen-sel" class="resumen-sel"></div>
      <label class="fecha-envio">Fecha de pago <input type="date" value="${P.fecha}" data-in="p-fecha"></label>
      <button class="btn primario" data-act="p-enviar" id="btn-enviar">Enviar a pagar</button>
    </div>
  </div>`;
}

function grupoProveedor(fs) {
  const f0 = fs[0];
  const prov = S.proveedor(f0.codigo);
  const t = S.totales(fs);
  return `<section class="grupo">
    <header class="grupo-cab">
      <div><h3>${esc(prov.nombre)}</h3><span class="sub">${esc(S.nombreSociedad(f0.sociedad))} · ${MONEDAS[f0.moneda].nombre} · ${fs.length} ${fs.length === 1 ? 'factura' : 'facturas'}</span></div>
      <div class="grupo-tot">${money(t[f0.moneda], f0.moneda)}
        <button class="btn chico" data-act="p-todas" data-v="${fs.map((f) => f.id).join(',')}">Marcar todas</button></div>
    </header>
    <div class="filas-fact">${fs.map(filaProgramar).join('')}</div>
  </section>`;
}

function filaProgramar(f) {
  const prog = S.programacionDe(f);
  const s = S.saldo(f);
  const sel = ui.prog.sel.has(f.id);
  if (prog) {
    return `<div class="fila-fact programada">
      <span class="ck"><input type="checkbox" disabled aria-label="Factura ${esc(f.documento)} ya programada"></span>
      <span class="doc">Fact. ${esc(f.documento)}</span>
      <span class="venc">${f.fechaVence ? fmtFecha(f.fechaVence) : ''} ${etiquetaDias(f)}</span>
      <span class="saldo">${money(s, f.moneda)}</span>
      <span class="apagar"><span class="estado prog">Programada ${money(prog.monto, f.moneda)} · ${fmtFecha(prog.fechaPago)}</span></span>
    </div>`;
  }
  return `<div class="fila-fact ${sel ? 'marcada' : ''}" data-fila="${f.id}">
    <label class="ck"><input type="checkbox" data-in="p-sel" data-v="${f.id}" ${sel ? 'checked' : ''}><span class="vh">Pagar factura ${esc(f.documento)}</span></label>
    <span class="doc">Fact. ${esc(f.documento)}${f.revision ? ' <span class="estado rev">Revisión</span>' : ''}${S.movimientosDe(f).length ? ` <span class="tag-abono" title="Ya tiene abonos">abonada</span>` : ''}</span>
    <span class="venc">${f.fechaVence ? fmtFecha(f.fechaVence) : ''} ${etiquetaDias(f)}</span>
    <span class="saldo">${money(s, f.moneda)}</span>
    <span class="apagar">${sel ? `<label><span class="vh">Monto a pagar de la factura ${esc(f.documento)}</span>
      <span class="input-moneda"><span>${MONEDAS[f.moneda].simbolo}</span><input type="text" inputmode="decimal" value="${esc(ui.prog.sel.get(f.id))}" data-in="p-monto" data-v="${f.id}"></span></label>
      <span class="tipo-pago" id="tipo-${f.id}"></span>` : ''}</span>
  </div>`;
}

/** Actualiza la barra inferior sin redibujar la pantalla (para no perder el foco al escribir). */
function actualizarBarra() {
  const $res = document.getElementById('resumen-sel');
  if (!$res) return;
  const tot = { CRC: 0, USD: 0 };
  let n = 0, errores = 0;
  for (const [id, txt] of ui.prog.sel) {
    const f = S.factura(id);
    if (!f) continue;
    n++;
    const m = parseNum(txt), s = S.saldo(f);
    const $tipo = document.getElementById('tipo-' + id);
    const $in = document.querySelector(`[data-in="p-monto"][data-v="${id}"]`);
    let tipo = '', malo = false;
    if (!(m > 0)) { tipo = 'Escriba el monto'; malo = true; }
    else if (m > s + 0.004) { tipo = 'Mayor que el saldo'; malo = true; }
    else if (m < s - 0.004) tipo = `Abono · quedan ${money(s - m, f.moneda)}`;
    else tipo = 'Total';
    if ($tipo) { $tipo.textContent = tipo; $tipo.className = 'tipo-pago ' + (malo ? 'malo' : tipo === 'Total' ? 'total' : 'abono'); }
    if ($in) $in.classList.toggle('malo', malo);
    if (malo) errores++; else tot[f.moneda] = round2(tot[f.moneda] + m);
  }
  $res.innerHTML = n
    ? `<strong>${n} ${n === 1 ? 'factura marcada' : 'facturas marcadas'}</strong><span>${totalesTxt(tot)}</span>${errores ? `<span class="malo">${errores} con monto por corregir</span>` : ''}`
    : '<span class="sub">Marque las facturas que se van a pagar</span>';
  document.getElementById('btn-enviar').disabled = !n || errores > 0;
  document.getElementById('barra-envio').classList.toggle('activa', n > 0);
}

// ---------------------------------------------------------------- Por aplicar (Jonathan y Daniel)

function vistaAplicar() {
  const lotes = S.lotesPendientes();
  const hechos = S.lotesAplicados().slice(0, 15);
  return `
  <div class="titulo-vista">
    <div>
      <h1>Por aplicar</h1>
      <p class="sub">Lo que Norman programó, agrupado por sociedad, moneda y fecha. Cada grupo genera un archivo para el banco.</p>
    </div>
  </div>
  ${lotes.length ? lotes.map(tarjetaLote).join('') : `<p class="vacio grande">No hay pagos programados pendientes. Cuando Norman programe pagos, aparecerán aquí.</p>`}

  ${hechos.length ? `<section class="bloque">
    <h2>Aplicados recientemente</h2>
    <div class="tabla-scroll"><table class="tabla">
      <thead><tr><th>Fecha</th><th>Sociedad</th><th>Comprobante</th><th class="num">Total</th><th>Proveedores</th><th>Aplicó</th></tr></thead>
      <tbody>${hechos.map((l) => `<tr><td>${fmtFecha(l.fecha)}</td><td>${esc(S.nombreSociedad(l.sociedad))}</td><td>${esc(l.comprobante)}</td>
        <td class="num">${money(l.total, l.moneda)}</td><td>${nLineas(l)}</td><td>${esc(l.aplicadoPor)}</td></tr>`).join('')}</tbody>
    </table></div>
  </section>` : ''}`;
}

function tarjetaLote(l) {
  const problemas = S.problemasLote(l);
  const nFact = l.lineas.reduce((s, x) => s + x.items.length, 0);
  const id = encodeURIComponent(l.clave);
  return `<section class="lote">
    <header class="lote-cab">
      <div>
        <span class="eyebrow">Pagar el ${fmtFecha(l.fecha)}</span>
        <h2>${esc(S.nombreSociedad(l.sociedad))} · ${MONEDAS[l.moneda].nombre}</h2>
        <span class="sub">${l.lineas.length} ${l.lineas.length === 1 ? 'proveedor' : 'proveedores'} · ${nFact} facturas · debita ${l.cuentaDebito ? `<span class="mono">${fmtIBAN(l.cuentaDebito)}</span>` : '<em>cuenta sin configurar</em>'}</span>
      </div>
      <div class="lote-total">${money(l.total, l.moneda)}</div>
    </header>
    <div class="tabla-scroll"><table class="tabla lote-tabla">
      <thead><tr><th>Proveedor</th><th>Facturas</th><th>IBAN</th><th class="num">Intereses / otros</th><th>Descripción para el banco</th><th class="num">Total</th></tr></thead>
      <tbody>${l.lineas.map((ln) => `<tr>
        <td><strong>${esc(ln.prov.nombre)}</strong></td>
        <td class="facts">${ln.items.map(({ p, f }) => `<div>${esc(f.documento)}: ${money(p.monto, f.moneda)}${p.monto < S.saldo(f) - 0.004 ? ' <span class="tag-abono">abono</span>' : ''}
            <button class="btn-link" data-act="a-devolver" data-v="${p.id}" title="Devolver a Norman">devolver</button></div>`).join('')}</td>
        <td class="mono">${ln.iban ? `${fmtIBAN(ln.iban)} ${errorIBAN(ln.iban) ? '<span class="malo">inválido</span>' : ''}` : `<a href="#/proveedores" data-act="ir-prov" data-v="${esc(ln.codigo)}" class="malo">Agregar IBAN</a>`}</td>
        <td class="num"><input class="mini" type="text" inputmode="decimal" value="${ln.extra ? fmtNum(ln.extra) : ''}" placeholder="0,00" data-in="a-extra" data-lote="${id}" data-v="${esc(ln.codigo)}" aria-label="Intereses u otros cargos para ${esc(ln.prov.nombre)}"></td>
        <td><input type="text" value="${esc(ln.descripcion)}" data-in="a-desc" data-lote="${id}" data-v="${esc(ln.codigo)}" aria-label="Descripción para ${esc(ln.prov.nombre)}"></td>
        <td class="num"><strong>${money(ln.total, l.moneda)}</strong></td></tr>`).join('')}</tbody>
    </table></div>
    ${problemas.length ? `<div class="aviso"><strong>Antes de generar el archivo:</strong><ul>${problemas.map((p) => `<li>${esc(p)}</li>`).join('')}</ul></div>` : ''}
    <div class="lote-acciones">
      <div class="paso">
        <span class="num-paso">1</span>
        <button class="btn oscuro" data-act="a-txt" data-v="${id}" ${problemas.length ? 'disabled' : ''}>Descargar archivo para el banco (.txt)</button>
        <button class="btn" data-act="a-ver-txt" data-v="${id}" ${problemas.length ? 'disabled' : ''}>Ver contenido</button>
      </div>
      <div class="paso aplicar">
        <span class="num-paso">2</span>
        <label>N° de comprobante <input type="text" id="comp-${id}" placeholder="Del banco"></label>
        <label>Fecha <input type="date" id="fecha-${id}" value="${l.fecha}"></label>
        <button class="btn primario" data-act="a-aplicar" data-v="${id}">Aplicar pago</button>
      </div>
    </div>
  </section>`;
}

function loteDe(idCodificado) {
  const clave = decodeURIComponent(idCodificado);
  return S.lotesPendientes().find((l) => l.clave === clave);
}

function textoArchivo(l) {
  const soc = S.sociedades()[l.sociedad];
  return construirArchivo({
    cliente: S.cfg('clienteBanco'),
    fecha: l.fecha,
    control: S.cfg('controlEncabezado'),
    debito: { iban: l.cuentaDebito, descripcion: soc?.descripcionBanco || ('PROVEE ' + S.nombreSociedad(l.sociedad)).toUpperCase() },
    creditos: l.lineas.map((ln) => ({ iban: ln.iban, monto: ln.total, descripcion: ln.descripcion })),
  });
}

function descargar(nombre, contenido, tipo) {
  const blob = contenido instanceof Blob ? contenido : new Blob([contenido], { type: tipo });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = nombre;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

// ---------------------------------------------------------------- Facturas

function vistaFacturas() {
  const F = ui.fact;
  const texto = F.q.trim().toLowerCase();
  let lista = S.listaFacturas().filter((f) => {
    if (F.soc && f.sociedad !== F.soc) return false;
    if (F.mon && f.moneda !== F.mon) return false;
    const e = S.estado(f);
    if (F.est === 'saldo' && e === 'Cancelada') return false;
    if (F.est && F.est !== 'saldo' && e !== F.est) return false;
    if (texto && !(`${S.proveedor(f.codigo).nombre} ${f.documento} ${S.movimientosDe(f).map((m) => m.comprobante).join(' ')}`.toLowerCase().includes(texto))) return false;
    return true;
  });
  lista.sort((a, b) => (b.fechaFactura || '').localeCompare(a.fechaFactura || ''));
  const total = lista.length;
  const t = S.totales(lista);
  lista = lista.slice(0, F.limite);
  return `
  <div class="titulo-vista">
    <div><h1>Facturas</h1><p class="sub">${S.listaFacturas().length} facturas en total</p></div>
    <div class="acciones"><button class="btn primario" data-act="f-nueva">+ Nueva factura</button></div>
  </div>
  <div class="filtros">
    <div class="fila-chips" role="group" aria-label="Sociedad">${chipsSociedad('f-soc', F.soc)}</div>
    <div class="fila-busqueda">
      <label class="buscar"><span class="vh">Buscar</span><input type="search" placeholder="Proveedor, factura o comprobante" value="${esc(F.q)}" data-in="f-q"></label>
      <label>Estado <select data-in="f-est">
        ${[['saldo', 'Con saldo'], ['', 'Todas'], ['Pendiente', 'Pendiente'], ['Programada', 'Programada'], ['Abonada', 'Abonada'], ['Revisión', 'Revisión'], ['Cancelada', 'Cancelada']]
          .map(([v, n]) => `<option value="${v}" ${F.est === v ? 'selected' : ''}>${n}</option>`).join('')}</select></label>
      <label>Moneda <select data-in="f-mon"><option value="">Ambas</option><option value="CRC" ${F.mon === 'CRC' ? 'selected' : ''}>Colones</option><option value="USD" ${F.mon === 'USD' ? 'selected' : ''}>Dólares</option></select></label>
    </div>
  </div>
  <p class="resumen-lista">${total} facturas · saldo ${totalesTxt(t)}</p>
  <div class="tabla-scroll"><table class="tabla clic">
    <thead><tr><th>Sociedad</th><th>Proveedor</th><th>Factura</th><th>Fecha</th><th>Vence</th><th class="num">Monto</th><th class="num">Saldo</th><th>Estado</th></tr></thead>
    <tbody>${lista.map((f) => `<tr data-act="f-abrir" data-v="${f.id}" tabindex="0">
      <td>${esc(S.nombreSociedad(f.sociedad))}</td><td>${esc(S.proveedor(f.codigo).nombre)}</td>
      <td><a href="#/facturas" data-act="f-abrir" data-v="${f.id}">${esc(f.documento)}</a></td>
      <td>${fmtFecha(f.fechaFactura)}</td>
      <td class="nowrap">${fmtFecha(f.fechaVence)} ${S.saldo(f) > 0.004 ? etiquetaDias(f) : ''}</td>
      <td class="num">${money(f.monto, f.moneda)}</td><td class="num"><strong>${money(S.saldo(f), f.moneda)}</strong></td>
      <td>${etiquetaEstado(S.estado(f))}</td></tr>`).join('')}</tbody>
  </table></div>
  ${total > lista.length ? `<p class="centro-txt"><button class="btn" data-act="f-mas">Mostrar más (${total - lista.length} restantes)</button></p>` : ''}`;
}

function dialogoFactura(id) {
  const f = S.factura(id);
  const prov = S.proveedor(f.codigo);
  const s = S.saldo(f);
  const tipos = { pago: 'Pago', abono: 'Abono', nc: 'Nota de crédito', cargo: 'Intereses / otros' };
  let acumulado = f.monto;
  abrirDialogo(`
    <header class="dlg-cab">
      <div><h2>Factura ${esc(f.documento)} ${etiquetaEstado(S.estado(f))}</h2>
        <p class="sub">${esc(prov.nombre)} · ${esc(S.nombreSociedad(f.sociedad))} · ${MONEDAS[f.moneda].nombre}</p></div>
      <button class="btn-cerrar" data-act="cerrar" aria-label="Cerrar">×</button>
    </header>
    <div class="dlg-cuerpo">
      <div class="cifras">
        <div><span>Monto</span><strong>${money(f.monto, f.moneda)}</strong></div>
        <div><span>Pagado</span><strong>${money(S.pagado(f), f.moneda)}</strong></div>
        <div><span>Saldo</span><strong>${money(s, f.moneda)}</strong></div>
      </div>
      <form class="form-grid" data-form="f-fechas" data-v="${f.id}">
        <label>Fecha de factura <input type="date" name="fechaFactura" value="${f.fechaFactura || ''}"></label>
        <label>Vencimiento ${f.venceEstimado ? '<span class="est">EST</span>' : ''}<input type="date" name="fechaVence" value="${f.fechaVence || ''}"></label>
        <label class="ancho">Observaciones <input type="text" name="observaciones" value="${esc(f.observaciones)}"></label>
        <label class="check ancho"><input type="checkbox" name="revision" ${f.revision ? 'checked' : ''}> En revisión</label>
        <div class="ancho"><button class="btn" type="submit">Guardar cambios</button></div>
      </form>
      <h3>Movimientos</h3>
      ${S.movimientosDe(f).length ? `<div class="tabla-scroll"><table class="tabla">
        <thead><tr><th>Fecha</th><th>Tipo</th><th>Comprobante</th><th class="num">Monto</th><th class="num">Saldo después</th><th>Registró</th></tr></thead>
        <tbody>${S.movimientosDe(f).map((m) => {
          if (m.tipo !== 'cargo') acumulado = round2(acumulado - m.monto);
          return `<tr><td>${fmtFecha(m.fecha) || '<span class="sub">sin fecha</span>'}</td><td>${tipos[m.tipo] || m.tipo}</td><td>${esc(m.comprobante)}</td>
            <td class="num">${money(m.monto, f.moneda)}</td><td class="num">${m.tipo === 'cargo' ? '' : money(acumulado, f.moneda)}</td><td>${esc(m.por)}</td></tr>`;
        }).join('')}</tbody></table></div>` : '<p class="sub">Sin movimientos.</p>'}
      ${s > 0.004 ? `
      <h3>Registrar abono o nota de crédito</h3>
      <form class="form-grid" data-form="f-mov" data-v="${f.id}">
        <label>Tipo <select name="tipo"><option value="abono">Abono / pago</option><option value="nc">Nota de crédito</option></select></label>
        <label>Monto <input type="text" name="monto" inputmode="decimal" value="${fmtNum(s)}"></label>
        <label>Fecha <input type="date" name="fecha" value="${hoyISO()}"></label>
        <label>N° comprobante o de nota <input type="text" name="comprobante"></label>
        <div class="ancho"><button class="btn primario" type="submit">Registrar</button></div>
      </form>` : ''}
    </div>`);
}

function dialogoNuevaFactura() {
  const socs = Object.keys(S.sociedades());
  const provs = Object.values(S.datos().proveedores).sort((a, b) => a.nombre.localeCompare(b.nombre));
  abrirDialogo(`
    <header class="dlg-cab"><h2>Nueva factura</h2><button class="btn-cerrar" data-act="cerrar" aria-label="Cerrar">×</button></header>
    <form class="dlg-cuerpo form-grid" data-form="f-crear">
      <label class="ancho">Sociedad <select name="sociedad" required>${socs.map((s) => `<option value="${esc(s)}">${esc(S.nombreSociedad(s))}</option>`).join('')}</select></label>
      <label class="ancho">Proveedor <input name="proveedor" list="lista-prov" required placeholder="Escriba para buscar">
        <datalist id="lista-prov">${provs.map((p) => `<option value="${esc(p.nombre)} · ${esc(p.codigo)}"></option>`).join('')}</datalist></label>
      <label>N° de documento <input name="documento" required></label>
      <label>Moneda <select name="moneda"><option value="CRC">Colones</option><option value="USD">Dólares</option></select></label>
      <label>Fecha de factura <input type="date" name="fechaFactura" value="${hoyISO()}" required data-in="nf-fecha"></label>
      <label>Plazo (días) <input type="number" name="plazo" value="30" min="0" data-in="nf-plazo"></label>
      <label>Vencimiento <input type="date" name="fechaVence" value="${sumarDias(hoyISO(), 30)}" id="nf-vence"></label>
      <label>Monto <input name="monto" inputmode="decimal" required></label>
      <label class="ancho">Observaciones <input name="observaciones"></label>
      <div class="ancho"><button class="btn primario" type="submit">Guardar factura</button></div>
    </form>`);
}

// ---------------------------------------------------------------- Proveedores

function vistaProveedores() {
  const P = ui.prov;
  const texto = P.q.trim().toLowerCase();
  const saldoPorProv = new Map();
  for (const f of S.conSaldo()) {
    const t = saldoPorProv.get(f.codigo) || { CRC: 0, USD: 0, n: 0 };
    t[f.moneda] = round2(t[f.moneda] + S.saldo(f)); t.n++;
    saldoPorProv.set(f.codigo, t);
  }
  let lista = Object.values(S.datos().proveedores).filter((p) => {
    if (texto && !(`${p.nombre} ${p.alias} ${p.codigo}`.toLowerCase().includes(texto))) return false;
    if (P.filtro === 'sin-iban' && (p.ibanCRC || p.ibanUSD)) return false;
    if (P.filtro === 'con-saldo' && !saldoPorProv.has(p.codigo)) return false;
    return true;
  }).sort((a, b) => a.nombre.localeCompare(b.nombre));
  const conIban = Object.values(S.datos().proveedores).filter((p) => p.ibanCRC || p.ibanUSD).length;
  return `
  <div class="titulo-vista">
    <div><h1>Proveedores</h1><p class="sub">${Object.keys(S.datos().proveedores).length} proveedores · ${conIban} con cuenta IBAN</p></div>
    <div class="acciones"><button class="btn primario" data-act="pv-nuevo">+ Nuevo proveedor</button></div>
  </div>
  <div class="filtros">
    <div class="fila-busqueda">
      <label class="buscar"><span class="vh">Buscar</span><input type="search" placeholder="Nombre, alias o código" value="${esc(P.q)}" data-in="pv-q"></label>
      <div class="fila-chips">${chip('pv-filtro', '', P.filtro, 'Todos')}${chip('pv-filtro', 'con-saldo', P.filtro, 'Con saldo')}${chip('pv-filtro', 'sin-iban', P.filtro, 'Sin IBAN')}</div>
    </div>
  </div>
  <div class="tabla-scroll"><table class="tabla clic">
    <thead><tr><th>Código</th><th>Proveedor</th><th>Alias banco</th><th>IBAN colones</th><th>IBAN dólares</th><th class="num">Saldo</th></tr></thead>
    <tbody>${lista.map((p) => {
      const t = saldoPorProv.get(p.codigo);
      return `<tr data-act="pv-abrir" data-v="${esc(p.codigo)}" tabindex="0">
      <td>${esc(p.codigo)}</td><td><a href="#/proveedores" data-act="pv-abrir" data-v="${esc(p.codigo)}">${esc(p.nombre)}</a></td><td>${esc(p.alias)}</td>
      <td class="mono">${p.ibanCRC ? fmtIBAN(p.ibanCRC) : '<span class="sub">—</span>'}</td>
      <td class="mono">${p.ibanUSD ? fmtIBAN(p.ibanUSD) : '<span class="sub">—</span>'}</td>
      <td class="num">${t ? totalesTxt(t) : ''}</td></tr>`;
    }).join('')}</tbody>
  </table></div>`;
}

function dialogoProveedor(codigo) {
  const nuevo = !codigo;
  const p = nuevo ? { codigo: S.siguienteCodigoProveedor(), nombre: '', alias: '', ibanCRC: '', ibanUSD: '', plazo: 30 } : S.proveedor(codigo);
  abrirDialogo(`
    <header class="dlg-cab"><h2>${nuevo ? 'Nuevo proveedor' : esc(p.nombre)}</h2><button class="btn-cerrar" data-act="cerrar" aria-label="Cerrar">×</button></header>
    <form class="dlg-cuerpo form-grid" data-form="pv-guardar" data-v="${esc(p.codigo)}">
      <label>Código <input name="codigo" value="${esc(p.codigo)}" ${nuevo ? '' : 'readonly'}></label>
      <label>Plazo habitual (días) <input type="number" name="plazo" min="0" value="${esc(p.plazo ?? 30)}"></label>
      <label class="ancho">Nombre <input name="nombre" value="${esc(p.nombre)}" required></label>
      <label class="ancho">Alias para la descripción del banco <input name="alias" value="${esc(p.alias)}" placeholder="Ej. Resusa, DAC"></label>
      <label class="ancho">IBAN en colones <input name="ibanCRC" class="mono" value="${esc(fmtIBAN(p.ibanCRC))}" placeholder="CR00 0000 0000 0000 0000 00" data-in="iban"><span class="ayuda" data-ayuda></span></label>
      <label class="ancho">IBAN en dólares <input name="ibanUSD" class="mono" value="${esc(fmtIBAN(p.ibanUSD))}" placeholder="CR00 0000 0000 0000 0000 00" data-in="iban"><span class="ayuda" data-ayuda></span></label>
      <div class="ancho"><button class="btn primario" type="submit">Guardar</button></div>
    </form>`);
  $dialog.querySelectorAll('[data-in="iban"]').forEach(validarCampoIBAN);
}

function validarCampoIBAN(input) {
  const ayuda = input.parentElement.querySelector('[data-ayuda]');
  const v = input.value.trim();
  const err = v ? errorIBAN(v) : null;
  input.classList.toggle('malo', !!err);
  if (ayuda) { ayuda.textContent = !v ? '' : err || 'IBAN válido'; ayuda.className = 'ayuda ' + (err ? 'malo' : 'bien'); }
  return !err;
}

// ---------------------------------------------------------------- Tablero

function vistaTablero() {
  const pend = S.conSaldo();
  const venc = pend.filter((f) => S.diasVencida(f) > 0);
  const semana = pend.filter((f) => { const d = S.diasVencida(f); return d != null && d <= 0 && d > -7; });
  const tramos = [['Por vencer', (d) => d <= 0], ['1–30 días', (d) => d >= 1 && d <= 30], ['31–60 días', (d) => d > 30 && d <= 60], ['61–90 días', (d) => d > 60 && d <= 90], ['Más de 90', (d) => d > 90]];
  const socs = Object.keys(S.sociedades());
  return `
  <div class="titulo-vista"><div><h1>Tablero</h1><p class="sub">Al ${fmtFecha(hoyISO())}</p></div></div>
  <div class="kpis">
    <div class="kpi"><span>Total por pagar</span><strong>${money(S.totales(pend).CRC, 'CRC')}</strong><strong class="usd">${money(S.totales(pend).USD, 'USD')}</strong><small>${pend.length} facturas</small></div>
    <div class="kpi alerta"><span>Vencido</span><strong>${money(S.totales(venc).CRC, 'CRC')}</strong><strong class="usd">${money(S.totales(venc).USD, 'USD')}</strong><small>${venc.length} facturas</small></div>
    <div class="kpi"><span>Vence en 7 días</span><strong>${money(S.totales(semana).CRC, 'CRC')}</strong><strong class="usd">${money(S.totales(semana).USD, 'USD')}</strong><small>${semana.length} facturas</small></div>
  </div>
  <div class="dos-col">
    <section class="bloque"><h2>Saldo por sociedad</h2>
      <table class="tabla"><thead><tr><th>Sociedad</th><th class="num">Facturas</th><th class="num">Colones</th><th class="num">Dólares</th></tr></thead>
      <tbody>${socs.map((s) => { const t = S.totales(pend.filter((f) => f.sociedad === s)); return `<tr><td>${esc(S.nombreSociedad(s))}</td><td class="num">${t.n}</td><td class="num">${money(t.CRC, 'CRC')}</td><td class="num">${money(t.USD, 'USD')}</td></tr>`; }).join('')}</tbody></table>
    </section>
    <section class="bloque"><h2>Antigüedad de saldos</h2>
      <table class="tabla"><thead><tr><th>Tramo</th><th class="num">Facturas</th><th class="num">Colones</th><th class="num">Dólares</th></tr></thead>
      <tbody>${tramos.map(([n, fn]) => { const t = S.totales(pend.filter((f) => { const d = S.diasVencida(f); return d != null && fn(d); })); return `<tr><td>${n}</td><td class="num">${t.n}</td><td class="num">${money(t.CRC, 'CRC')}</td><td class="num">${money(t.USD, 'USD')}</td></tr>`; }).join('')}</tbody></table>
      <p class="sub">EST: ${pend.filter((f) => f.venceEstimado).length} facturas con saldo usan vencimiento estimado (factura + 30 días).</p>
    </section>
  </div>`;
}

// ---------------------------------------------------------------- Configuración

function vistaConfig() {
  const socs = Object.entries(S.sociedades());
  const e = Sync.estado;
  const imp = S.cfg('importado');
  return `
  <div class="titulo-vista"><div><h1>Configuración</h1><p class="sub">${imp ? 'Excel cargado: ' + esc(imp) : 'Aún no se ha cargado el Excel'}</p></div></div>

  <section class="bloque">
    <h2>Google Sheets</h2>
    <p class="sub">Toda la información se guarda en la hoja de Google. Cada computadora tiene además una copia para trabajar sin internet.</p>
    <div class="acciones-fila">
      ${e.hoja ? `<a class="btn" href="${esc(e.hoja)}" target="_blank" rel="noopener">Abrir la hoja de Google</a>` : ''}
      <button class="btn" data-act="sync-ahora">Sincronizar ahora</button>
      <button class="btn" data-act="cfg-excel" ${S.hayDatos() ? '' : 'disabled'}>Exportar facturas a Excel</button>
      ${S.esAdmin() ? '<button class="btn" data-act="conexion">Cambiar conexión</button>' : ''}
    </div>
    <p class="sub">${e.pendientes.length ? `${e.pendientes.length} cambios esperando conexión para enviarse.` : 'No hay cambios pendientes de enviar.'}${e.error ? ' Último error: ' + esc(e.error) : ''}</p>
  </section>

  <section class="bloque">
    <h2>Banco</h2>
    <form class="form-grid" data-form="cfg-banco">
      <label>Código de cliente en el banco <input name="clienteBanco" value="${esc(S.cfg('clienteBanco'))}" placeholder="Ej. 1234"></label>
      <label>Campo de control del encabezado (posiciones 27-38) <input name="controlEncabezado" value="${esc(S.cfg('controlEncabezado'))}" placeholder="12 dígitos, por confirmar con el banco"></label>
      <p class="ancho sub">El campo de control aún no está confirmado: pregunte al banco qué valor lleva. Si se deja vacío, se llena con ceros.</p>
      <div class="ancho"><button class="btn primario" type="submit">Guardar</button></div>
    </form>
  </section>

  <section class="bloque">
    <h2>Cuentas de las sociedades (cuenta que se debita)</h2>
    ${socs.length ? socs.map(([k, s]) => `
    <form class="form-grid sociedad-form" data-form="cfg-soc" data-v="${esc(k)}">
      <h3 class="ancho">${esc(k)}</h3>
      <label>Nombre corto <input name="corto" value="${esc(s.corto)}"></label>
      <label>Descripción en la línea de débito <input name="descripcionBanco" value="${esc(s.descripcionBanco)}" placeholder="Ej. PROVEE ${esc(String(s.corto).toUpperCase())}"></label>
      <label>IBAN colones <input name="cuentaCRC" class="mono" value="${esc(fmtIBAN(s.cuentaCRC))}" data-in="iban"><span class="ayuda" data-ayuda></span></label>
      <label>IBAN dólares <input name="cuentaUSD" class="mono" value="${esc(fmtIBAN(s.cuentaUSD))}" data-in="iban"><span class="ayuda" data-ayuda></span></label>
      <div class="ancho"><button class="btn" type="submit">Guardar ${esc(s.corto)}</button></div>
    </form>`).join('') : '<p class="sub">Aparecerán al cargar el Excel.</p>'}
  </section>

  ${S.esAdmin() ? `<section class="bloque">
    <h2>Volver a cargar el Excel</h2>
    <p class="sub">Reemplaza en Google Sheets todas las facturas, pagos y programaciones por los del Excel. Se conservan IBAN, alias y cuentas. Úselo solo para empezar de nuevo.</p>
    <label class="btn">Elegir archivo de Excel<input type="file" accept=".xlsx,.xls" data-in="importar" hidden></label>
    <p class="nota" id="progreso-import"></p>
  </section>` : ''}`;
}

// ---------------------------------------------------------------- Diálogo

function abrirDialogo(html) {
  $dialog.innerHTML = html;
  if (!$dialog.open) $dialog.showModal();
}
function cerrarDialogo() { if ($dialog.open) $dialog.close(); }
$dialog.addEventListener('click', (e) => { if (e.target === $dialog) cerrarDialogo(); });

function avisar(msg, tipo = 'ok') {
  const t = document.createElement('div');
  t.className = 'toast ' + tipo;
  t.setAttribute('role', 'status');
  t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), tipo === 'error' ? 6000 : 3500);
}

// ---------------------------------------------------------------- Eventos

function despuesDeRender(ruta) {
  if (ruta === 'programar') actualizarBarra();
  if (ruta === 'config') document.querySelectorAll('[data-in="iban"]').forEach(validarCampoIBAN);
}

document.addEventListener('click', (e) => {
  const el = e.target.closest('[data-act]');
  if (!el) return;
  const act = el.dataset.act, v = el.dataset.v;
  if (el.tagName === 'A' || el.tagName === 'TR') e.preventDefault();
  acciones[act]?.(v, el, e);
});

document.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && e.target.matches('tr[data-act]')) acciones[e.target.dataset.act]?.(e.target.dataset.v);
});

const acciones = {
  'sync-ahora': () => Sync.sincronizar(),
  conexion: () => dialogoConexion(),
  cuenta: () => dialogoCuenta(),
  salir: async () => {
    const n = Sync.estado.pendientes.length;
    if (!confirm(n ? `Hay ${n} cambios que aún no llegan a Google Sheets y se perderán. ¿Cerrar sesión de todos modos?` : '¿Cerrar sesión en esta computadora?')) return;
    cerrarDialogo();
    await Sync.cerrarSesion();
    location.hash = '';
  },
  cerrar: cerrarDialogo,

  'p-soc': (v) => { ui.prog.soc = v; render(); },
  'p-mon': (v) => { ui.prog.mon = v; render(); },
  'p-todas': (v) => { for (const id of v.split(',')) { const f = S.factura(id); if (!S.programacionDe(f) && !ui.prog.sel.has(id)) ui.prog.sel.set(id, fmtNum(S.saldo(f))); } render(); },
  'p-quitar': (v) => { if (confirm('¿Quitar este pago de lo programado?')) resultado(S.quitarProgramacion(v, 'Quitado por Norman'), 'Pago quitado'); },
  'p-enviar': () => {
    const items = [...ui.prog.sel].map(([facturaId, txt]) => ({ facturaId, monto: parseNum(txt) }));
    const r = S.programarPagos(items, ui.prog.fecha);
    if (r.error) { avisar(r.error, 'error'); return; }
    ui.prog.sel.clear();
    render();
    avisar(`${items.length} ${items.length === 1 ? 'pago programado' : 'pagos programados'} para el ${fmtFecha(ui.prog.fecha)}` + (r.avisos.length ? ' · ' + r.avisos.join(' · ') : ''));
  },

  'a-devolver': (v) => { if (confirm('¿Devolver esta factura a Norman? Saldrá de este lote.')) resultado(S.quitarProgramacion(v, 'Devuelto'), 'Devuelta a Norman'); },
  'a-txt': (v) => {
    const l = loteDe(v);
    try {
      const txt = textoArchivo(l);
      const soc = S.nombreSociedad(l.sociedad).toUpperCase().replace(/[^A-Z0-9]+/g, '');
      descargar(`PROVEEDORES_${soc}_${l.moneda}_${l.fecha.split('-').reverse().join('')}.txt`, txt, 'text/plain');
    } catch (err) { avisar(err.message, 'error'); }
  },
  'a-ver-txt': (v) => {
    const l = loteDe(v);
    try {
      const txt = textoArchivo(l);
      abrirDialogo(`<header class="dlg-cab"><h2>Contenido del archivo</h2><button class="btn-cerrar" data-act="cerrar" aria-label="Cerrar">×</button></header>
        <div class="dlg-cuerpo"><pre class="txt">${esc(txt.replace(/ +\r\n/g, '\r\n'))}</pre><p class="sub">Los espacios al final de cada línea no se muestran; el archivo descargado sí los lleva (214 caracteres por línea).</p></div>`);
    } catch (err) { avisar(err.message, 'error'); }
  },
  'a-aplicar': (v) => {
    const l = loteDe(v);
    const comp = document.getElementById('comp-' + v).value.trim();
    const fecha = document.getElementById('fecha-' + v).value;
    if (!confirm(`¿Aplicar el pago de ${money(l.total, l.moneda)} con el comprobante ${comp || '(vacío)'}?`)) return;
    resultado(S.aplicarLote(l, { comprobante: comp, fecha }), 'Pago aplicado');
  },
  'ir-prov': (v) => { location.hash = '#/proveedores'; setTimeout(() => dialogoProveedor(v), 0); },

  'f-soc': (v) => { ui.fact.soc = v; ui.fact.limite = 150; render(); },
  'f-mas': () => { ui.fact.limite += 300; render(); },
  'f-abrir': (v) => dialogoFactura(v),
  'f-nueva': () => dialogoNuevaFactura(),

  'pv-filtro': (v) => { ui.prov.filtro = v; render(); },
  'pv-abrir': (v) => dialogoProveedor(v),
  'pv-nuevo': () => dialogoProveedor(null),

  'cfg-excel': () => exportarExcel(),
};

let tBusqueda;
document.addEventListener('input', (e) => {
  const el = e.target, k = el.dataset.in;
  if (!k) return;
  if (k === 'p-monto') { ui.prog.sel.set(el.dataset.v, el.value); actualizarBarra(); return; }
  if (k === 'iban') { validarCampoIBAN(el); return; }
  if (k === 'nf-fecha' || k === 'nf-plazo') {
    const form = el.form;
    const ff = form.fechaFactura.value, pl = Number(form.plazo.value) || 0;
    if (ff) form.fechaVence.value = sumarDias(ff, pl);
    return;
  }
  const busquedas = { 'p-q': (x) => { ui.prog.q = x; }, 'f-q': (x) => { ui.fact.q = x; ui.fact.limite = 150; }, 'pv-q': (x) => { ui.prov.q = x; } };
  if (busquedas[k]) {
    busquedas[k](el.value);
    clearTimeout(tBusqueda);
    tBusqueda = setTimeout(() => {
      const pos = el.selectionStart;
      render();
      const nuevo = document.querySelector(`[data-in="${k}"]`);
      if (nuevo) { nuevo.focus(); nuevo.setSelectionRange(pos, pos); }
    }, 250);
  }
});

document.addEventListener('change', async (e) => {
  const el = e.target, k = el.dataset.in;
  if (!k) return;
  if (k === 'p-sel') {
    const f = S.factura(el.dataset.v);
    if (el.checked) ui.prog.sel.set(f.id, fmtNum(S.saldo(f))); else ui.prog.sel.delete(f.id);
    render();
    if (el.checked) document.querySelector(`[data-in="p-monto"][data-v="${f.id}"]`)?.select();
    return;
  }
  if (k === 'p-venc') { ui.prog.vencidas = el.checked; render(); return; }
  if (k === 'p-fecha') { ui.prog.fecha = el.value; return; }
  if (k === 'f-est') { ui.fact.est = el.value; ui.fact.limite = 150; render(); return; }
  if (k === 'f-mon') { ui.fact.mon = el.value; render(); return; }
  if (k === 'a-extra' || k === 'a-desc') {
    const clave = decodeURIComponent(el.dataset.lote);
    if (k === 'a-extra') {
      const n = el.value.trim() ? parseNum(el.value) : 0;
      if (!(n >= 0)) { avisar('Monto de intereses inválido', 'error'); return; }
      S.guardarBorrador(clave, el.dataset.v, { extra: round2(n), descripcion: undefined });
    } else S.guardarBorrador(clave, el.dataset.v, { descripcion: el.value });
    render();
    return;
  }
  if (k === 'importar') {
    const file = el.files[0];
    el.value = '';
    if (!file) return;
    if (S.hayDatos() && !confirm('Esto reemplaza en Google Sheets todas las facturas, pagos y programaciones por los del Excel. Se conservan IBAN, alias y cuentas. ¿Continuar?')) return;
    const $p = document.getElementById('progreso-import');
    const paso = (t) => { if ($p) $p.textContent = t; };
    try {
      paso('Leyendo el Excel…');
      const res = importarExcel(await file.arrayBuffer(), window.XLSX);
      paso(`Enviando ${res.facturas.length} facturas a Google Sheets… (puede tardar un minuto)`);
      await Sync.importar(S.paqueteImportacion(res));
      paso('');
      avisar(`Se cargaron ${res.facturas.length} facturas en Google Sheets` + (res.avisos.length ? ` (${res.avisos.length} avisos)` : ''));
    } catch (err) { paso(''); avisar('No se pudo cargar el Excel: ' + err.message, 'error'); }
  }
});

document.addEventListener('submit', (e) => {
  const form = e.target, k = form.dataset.form;
  if (!k) return;
  e.preventDefault();
  const d = Object.fromEntries(new FormData(form));
  if (k === 'login') {
    const btn = form.querySelector('button[type=submit]');
    btn.disabled = true;
    btn.textContent = 'Entrando…';
    Sync.iniciarSesion(d.usuario, d.contrasena).then(() => {
      location.hash = S.esNorman() ? '#/programar' : '#/aplicar';
      render();
    }).catch((err) => {
      avisar(navigator.onLine ? err.message : 'Para entrar la primera vez se necesita internet', 'error');
      btn.disabled = false;
      btn.textContent = 'Entrar';
    });
  } else if (k === 'conexion') {
    const btn = form.querySelector('button[type=submit]');
    btn.disabled = true;
    btn.textContent = 'Conectando…';
    Sync.configurarConexion(d.url, d.clave).then(() => {
      cerrarDialogo();
      avisar('Computadora conectada con Google Sheets');
      render();
    }).catch((err) => {
      avisar(navigator.onLine ? err.message : 'Para conectar la computadora se necesita internet', 'error');
      btn.disabled = false;
      btn.textContent = 'Conectar';
    });
  } else if (k === 'nueva-clave' || k === 'cambiar-clave') {
    if (d.nueva !== d.repetir) { avisar('Las dos contraseñas nuevas no coinciden', 'error'); return; }
    const btn = form.querySelector('button[type=submit]');
    btn.disabled = true;
    Sync.cambiarContrasena(d.actual || '', d.nueva).then(() => {
      cerrarDialogo();
      avisar('Contraseña guardada');
      render();
    }).catch((err) => { avisar(navigator.onLine ? err.message : 'Para cambiar la contraseña se necesita internet', 'error'); btn.disabled = false; });
  } else if (k === 'f-fechas') {
    const f = S.factura(form.dataset.v);
    const cambios = { fechaFactura: d.fechaFactura || '', observaciones: d.observaciones, revision: !!d.revision, fechaVence: d.fechaVence || '' };
    if (resultado(S.actualizarFactura(f.id, cambios), 'Cambios guardados')) dialogoFactura(f.id);
  } else if (k === 'f-mov') {
    if (resultado(S.registrarMovimiento(form.dataset.v, { tipo: d.tipo, monto: parseNum(d.monto), fecha: d.fecha, comprobante: d.comprobante }), 'Movimiento registrado')) dialogoFactura(form.dataset.v);
  } else if (k === 'f-crear') {
    const codigo = (d.proveedor.match(/·\s*(\S+)$/) || [])[1];
    if (!codigo || !S.datos().proveedores[codigo]) { avisar('Elija el proveedor de la lista (o créelo en Proveedores)', 'error'); return; }
    const monto = parseNum(d.monto);
    if (!(monto > 0)) { avisar('Monto inválido', 'error'); return; }
    const r = S.nuevaFactura({ sociedad: d.sociedad, codigo, documento: d.documento.trim(), moneda: d.moneda, fechaFactura: d.fechaFactura, fechaVence: d.fechaVence || sumarDias(d.fechaFactura, 30), monto, observaciones: d.observaciones });
    if (resultado(r, 'Factura guardada')) cerrarDialogo();
  } else if (k === 'pv-guardar') {
    const nuevo = !form.querySelector('[name=codigo]').readOnly;
    if (resultado(S.guardarProveedor(d.codigo.trim(), { nombre: d.nombre, alias: d.alias, ibanCRC: d.ibanCRC, ibanUSD: d.ibanUSD, plazo: d.plazo }, nuevo), 'Proveedor guardado')) cerrarDialogo();
  } else if (k === 'cfg-banco') {
    resultado(S.guardarConfig({ clienteBanco: d.clienteBanco.trim(), controlEncabezado: d.controlEncabezado.trim() }), 'Configuración guardada');
  } else if (k === 'cfg-soc') {
    resultado(S.guardarSociedad(form.dataset.v, d), 'Sociedad guardada');
  }
});

function exportarExcel() {
  const filas = S.listaFacturas().map((f) => {
    const ultimo = [...S.movimientosDe(f)].reverse().find((m) => m.tipo !== 'cargo');
    return {
      'Origen de la Factura': f.sociedad, 'Código': f.codigo, 'Nombre del Proveedor': S.proveedor(f.codigo).nombre,
      'Fec/Factura': fmtFecha(f.fechaFactura), 'Fec/Vence': fmtFecha(f.fechaVence), 'Vence estimado': f.venceEstimado ? 'Sí' : '',
      'Número Documento': f.documento, 'Moneda': f.moneda === 'USD' ? 'USD' : 'COL', 'Monto': f.monto, 'Pagado': S.pagado(f), 'Saldo': S.saldo(f),
      'Estado': S.estado(f), 'Último pago': fmtFecha(ultimo?.fecha), 'N° Comprobante': ultimo?.comprobante || '', 'Observaciones': f.observaciones,
    };
  });
  const XLSX = window.XLSX;
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(filas), 'Facturas');
  XLSX.writeFile(wb, `cuentas-por-pagar-${hoyISO()}.xlsx`);
}

/** Muestra el resultado de un cambio. Devuelve true si se aplicó. */
function resultado(r, okMsg) {
  if (r.error) { avisar(r.error, 'error'); return false; }
  avisar(okMsg + (r.avisos?.length ? ' · ' + r.avisos.join(' · ') : ''));
  return true;
}

Sync.alAvisar((msg, tipo) => avisar(msg, tipo === 'aviso' ? 'ok' : tipo));
Sync.iniciar();
render();

if ('serviceWorker' in navigator && location.protocol === 'https:') {
  navigator.serviceWorker.register('sw.js').catch((e) => console.warn('Sin modo sin internet:', e));
}
