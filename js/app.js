// Pantallas y eventos de la aplicación.
import * as S from './store.js';
import * as Sync from './sync.js';
import { VERSION, USUARIOS, API_URL } from './config.js';
import { importarExcel } from './importer.js';
import { crearPdfReporte, crearPdfEstadoCuenta } from './reporte-pdf.js';
import { crearAvisoPago } from './aviso-pdf.js';
import { CAMPOS as CAMPOS_IMP, leerFilas, filaEncabezado, sugerirMapeo, prepararFacturas, detectarGTI } from './importar-facturas.js';
import { construirArchivo } from './banco.js';
import {
  money, fmtNum, parseNum, parseMonto, formatearMontoEscrito, fechaHora, fmtFecha, hoyISO, esc, errorIBAN, fmtIBAN, sumarDias, round2, MONEDAS,
} from './util.js';

const $app = document.getElementById('app');
const $nav = document.getElementById('nav');
const $user = document.getElementById('usuario');
const $dialog = document.getElementById('dialogo');

// Estado temporal de pantallas (no se guarda)
const ui = {
  login: '',
  prog: { soc: '', mon: '', q: '', vencidas: false, sel: new Map(), fecha: '', verProgramado: true },
  fact: { soc: '', mon: '', est: 'saldo', q: '', limite: 150 },
  canc: { soc: '', mon: '', q: '', limite: 100 },
  prov: { q: '', filtro: '', abierto: null },
  rep: { tipo: 'cxp', soc: '', prov: '', mon: 'CRC', desde: '', hasta: '' },
  ec: { prov: '', soc: '', mon: 'CRC', alcance: 'saldo', desde: '', hasta: '' },
  provDrop: null,
  forma: {},      // Por aplicar: 'banco' | 'manual' por proveedor de cada lote
  adj: {},        // archivos elegidos para adjuntar, antes de subirlos
  apAbierto: {},  // Aplicados recientemente: lotes desplegados
};

// ---------------------------------------------------------------- Navegación

const RUTAS = {
  programar: { titulo: 'Programar pagos', vista: vistaProgramar },
  aplicar: { titulo: 'Por aplicar', vista: vistaAplicar },
  facturas: { titulo: 'Facturas', vista: vistaFacturas },
  canceladas: { titulo: 'Canceladas', vista: vistaCanceladas },
  proveedores: { titulo: 'Proveedores', vista: vistaProveedores },
  tablero: { titulo: 'Tablero', vista: vistaTablero },
  reportes: { titulo: 'Reportes', vista: vistaReportes },
  config: { titulo: 'Configuración', vista: vistaConfig },
};

function menuPara(usuario) {
  if (usuario?.rol === 'aprueba') return ['tablero', 'programar', 'facturas', 'canceladas', 'reportes'];
  return ['tablero', 'aplicar', 'facturas', 'canceladas', 'proveedores', 'reportes', 'config']; // Programar pagos: solo Norman
}

function rutaActual() {
  const r = location.hash.replace(/^#\/?/, '').split('?')[0];
  if (RUTAS[r]) return r;
  return 'tablero';
}

const COLORES_USUARIO = ['u-1', 'u-2', 'u-3', 'u-4', 'u-5', 'u-6'];
function colorUsuario(usuario) {
  const lista = (Sync.usuariosConocidos() || USUARIOS).map((x) => x.usuario);
  let i = lista.indexOf(usuario);
  if (i < 0) i = [...String(usuario)].reduce((a, c) => a + c.charCodeAt(0), 0);
  return COLORES_USUARIO[i % COLORES_USUARIO.length];
}

const ROLES = { admin: 'Administra y aplica', aplica: 'Aplica pagos', aprueba: 'Aprueba pagos' };

function indicadorRed() {
  return indicadorConexion();
}

function indicadorConexion() {
  const e = Sync.estado, n = e.pendientes.length;
  const pend = n ? ` · ${n} ${n === 1 ? 'cambio por enviar' : 'cambios por enviar'}` : '';
  if (e.red === 'offline') return `<span class="red off" title="Los cambios se guardan aquí y se envían al volver la conexión">Sin internet${pend}</span>`;
  if (e.red === 'sync') return `<span class="red sync">Sincronizando…${pend}</span>`;
  if (e.red === 'error') return `<button class="red err" data-act="sync-ahora" title="${esc(e.error)}">No se pudo sincronizar${pend} · reintentar</button>`;
  return `<button class="red ok" data-act="sync-ahora" title="Sincronizar ahora">Sincronizado${e.ultimoSync ? ' ' + fechaHora(e.ultimoSync) : ''}${pend}</button>`;
}

function render() {
  const u = S.usuarioActual();
  let ruta = rutaActual();
  if (S.usuarioActual() && !menuPara(S.usuarioActual()).includes(ruta) && ruta !== 'tablero') ruta = 'tablero';
  const foco = guardarFoco();
  $nav.innerHTML = u ? menuPara(u).map((r) => {
    const n = r === 'aplicar' ? S.programacionesPendientes().filter((p) => !S.factura(p.facturaId)?.revision).length + porRevisar().length : 0;
    return `<a href="#/${r}" class="${r === ruta ? 'activo' : ''}" ${r === ruta ? 'aria-current="page"' : ''}>${RUTAS[r].titulo}${n ? ` <span class="burbuja">${n}</span>` : ''}</a>`;
  }).join('') : '';
  $user.innerHTML = u
    ? `${indicadorRed()}<button class="usuario-info" data-act="cambiar-usuario" title="Cambiar de usuario"><span class="avatar ${colorUsuario(u.usuario)}" aria-hidden="true">${esc(u.nombre.split(' ').map((p) => p[0]).join('').slice(0, 2))}</span><span><strong>${esc(u.nombre)}</strong><small>${esc(ROLES[u.rol] || u.rol)}</small></span></button>`
    : '';
  document.title = `${RUTAS[ruta].titulo} · Cuentas por Pagar`;

  if (!Sync.estado.cargado) { $app.innerHTML = '<p class="vacio">Cargando…</p>'; return; }
  if (!u) { $app.innerHTML = vistaLogin(); return; }
  if (u.debeCambiar) { $nav.innerHTML = ''; $app.innerHTML = vistaLogin(); return; }
  if (!S.hayDatos() && ruta !== 'config') { $app.innerHTML = Sync.estado.desde ? vistaBienvenida() : vistaCargando(); return; }
  const html = RUTAS[ruta].vista();
  if (ruta === rutaDibujada && $app.firstChild) {
    // Misma pantalla: solo se cambia lo que cambió. El campo donde se escribe no se toca,
    // así el teclado del celular no se reinicia con cada búsqueda.
    const t = document.createElement('template');
    t.innerHTML = html;
    etiquetarTablas(t.content);
    const ediciones = capturarEdiciones();
    parchar($app, t.content);
    restaurarEdiciones(ediciones);
  } else {
    $app.innerHTML = html;
  }
  rutaDibujada = ruta;
  despuesDeRender(ruta);
  restaurarFoco(foco);
}
let rutaDibujada = null;

/** Actualiza `viejo` para que quede igual a `nuevo`, sin reemplazar el elemento que tiene el foco. */
function parchar(viejo, nuevo) {
  const activo = document.activeElement;
  const sel = activo && activo !== document.body && viejo.contains(activo) ? selectorDe(activo) : null;
  const a = [...viejo.childNodes], b = [...nuevo.childNodes];
  const ia = sel ? a.findIndex((x) => x.contains(activo)) : -1;
  const jb = ia < 0 ? -1 : b.findIndex((y) => y.nodeType === 1 && y.tagName === a[ia].tagName && (y.matches(sel) || y.querySelector(sel)));
  if (jb < 0) { viejo.replaceChildren(...b); return; }
  const x = a[ia], y = b[jb];
  a.forEach((n) => { if (n !== x) n.remove(); });
  x.before(...b.slice(0, jb));
  x.after(...b.slice(jb + 1));
  copiarAtributos(x, y, x === activo);
  if (x !== activo) parchar(x, y);
}
/** Selector único de un campo: por id, por data-in (+data-v / data-k / data-lote) o por formulario + nombre. */
function selectorDe(el) {
  if (el.id) return '#' + CSS.escape(el.id);
  const f = el.form;
  if (el.name && f?.dataset.form) return `form[data-form="${CSS.escape(f.dataset.form)}"]${f.dataset.v != null ? `[data-v="${CSS.escape(f.dataset.v)}"]` : ''} [name="${CSS.escape(el.name)}"]`;
  if (el.dataset.in) {
    return `[data-in="${CSS.escape(el.dataset.in)}"]` + ['v', 'k', 'lote'].map((a) => el.dataset[a] != null ? `[data-${a}="${CSS.escape(el.dataset[a])}"]` : '').join('');
  }
  return null;
}

/**
 * Lo escrito y aún no guardado (el valor ya no es el que vino de los datos). Al redibujar por una
 * sincronización se vuelve a poner, para que no se borre. Si el dato guardado cambió, gana el nuevo.
 */
function capturarEdiciones() {
  const out = [];
  for (const el of $app.querySelectorAll('input, select, textarea')) {
    if (el.type === 'file' || el.type === 'hidden') continue;
    const sucio = el.type === 'checkbox' || el.type === 'radio' ? el.checked !== el.defaultChecked
      : el.tagName === 'SELECT' ? [...el.options].some((o) => o.selected !== o.defaultSelected)
      : el.value !== el.defaultValue;
    if (!sucio) continue;
    const sel = selectorDe(el);
    if (sel) out.push({ sel, valor: el.value, checked: el.checked, base: el.tagName === 'SELECT' ? null : el.defaultValue });
  }
  return out;
}
function restaurarEdiciones(lista) {
  for (const x of lista) {
    let el;
    try { el = $app.querySelector(x.sel); } catch { continue; }
    if (!el || el === document.activeElement) continue;
    if (x.base != null && el.defaultValue !== x.base) continue; // el dato cambió mientras tanto
    if (el.type === 'checkbox' || el.type === 'radio') el.checked = x.checked; else el.value = x.valor;
  }
}
function copiarAtributos(x, y, esActivo) {
  for (const { name } of [...x.attributes]) if (!y.hasAttribute(name)) x.removeAttribute(name);
  for (const { name, value } of [...y.attributes]) {
    if (esActivo && name === 'value') continue;
    if (x.getAttribute(name) !== value) x.setAttribute(name, value);
  }
}

/** Al llegar cambios de otra persona se redibuja la pantalla: se conserva el campo donde se escribe. */
function guardarFoco() {
  const el = document.activeElement;
  if (!el || !$app.contains(el) || !el.matches('input, select, textarea')) return null;
  const sel = selectorDe(el);
  let ini = null, fin = null;
  try { ini = el.selectionStart; fin = el.selectionEnd; } catch { /* campo sin selección */ }
  return { sel, valor: el.value, ini, fin };
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

const PIN = 'type="password" inputmode="numeric" pattern="[0-9]{4}" maxlength="4" autocomplete="off"';
const iniciales = (n) => esc(n.split(' ').map((p) => p[0]).join('').slice(0, 2));

/** Pantalla de inicio: las personas que usan la aplicación. */
function vistaLogin() {
  const lista = (Sync.usuariosConocidos() || USUARIOS).filter((x) => x.activo !== false);
  const elegido = lista.find((x) => x.usuario === ui.login);
  if (elegido) return vistaClave(elegido);
  return `<section class="centro">
    <img class="logo-inicio" src="img/logo.png" alt="Beneficiadora Santa Eduviges">
    <h1>¿Quién está usando la aplicación?</h1>
    <p class="sub">Toque su nombre para entrar.</p>
    <div class="usuarios">${lista.map((u) => `
      <button class="tarjeta-usuario" data-act="elegir" data-v="${esc(u.usuario)}">
        <span class="avatar grande ${colorUsuario(u.usuario)}" aria-hidden="true">${iniciales(u.nombre)}</span>
        <strong>${esc(u.nombre)}</strong><small>${esc(ROLES[u.rol] || u.rol)}</small>
      </button>`).join('')}</div>
  </section>`;
}

/** Solo la primera vez en un dispositivo: la contraseña del servicio. */
function vistaClave(u) {
  return `<section class="centro">
    <span class="avatar grande ${colorUsuario(u.usuario)}" aria-hidden="true">${iniciales(u.nombre)}</span>
    <h1>${esc(u.nombre)}</h1>
    <p class="sub">Primera vez en este dispositivo: escriba la contraseña del servicio. Solo se pide una vez; después se entra tocando el nombre.</p>
    <form class="form-login" data-form="login">
      <input type="hidden" name="usuario" value="${esc(u.usuario)}">
      ${API_URL ? '' : '<label>Dirección del servicio <input name="url" type="url" required placeholder="https://script.google.com/macros/s/…/exec" autocomplete="off"></label>'}
      <label>Contraseña del servicio <input name="servicio" type="password" required autocomplete="off" autofocus>
        <span class="ayuda">La tiene Jonathan.</span></label>
      <button class="btn primario grande" type="submit">Entrar</button>
    </form>
    <p class="nota"><button class="btn-link" data-act="elegir" data-v="">← Elegir otro usuario</button></p>
  </section>`;
}

function formConexion(url = '') {
  return `<form class="form-login" data-form="conexion">
      <label>Dirección del servicio <input name="url" type="url" required value="${esc(url)}" placeholder="https://script.google.com/macros/s/…/exec" autocomplete="off"></label>
      <label>Contraseña del servicio <input name="clave" type="password" required autocomplete="off"></label>
      <button class="btn primario" type="submit">Guardar conexión</button>
    </form>`;
}

/** La conexión solo se muestra después de escribir la contraseña del servicio. */
function dialogoConexion(desbloqueada = false) {
  const c = Sync.conexion();
  const cuerpo = !desbloqueada && c
    ? `<form class="form-login" data-form="ver-conexion">
        <label>Contraseña del servicio <input name="clave" type="password" required autocomplete="off" autofocus></label>
        <button class="btn primario" type="submit">Ver conexión</button>
      </form>`
    : `${c ? `<p class="sub">Dirección actual:</p><p class="mono caja-url">${esc(c.url)}</p>
        <p class="sub">Para cambiarla, escriba la nueva dirección y la contraseña del servicio.</p>` : ''}
      ${formConexion(c?.url || API_URL)}`;
  abrirDialogo(`
    <header class="dlg-cab"><div><h2>Conexión con Google Sheets</h2>
      <p class="sub">Dirección del servicio y contraseña del servicio de este dispositivo.</p></div>
      <button class="btn-cerrar" data-act="cerrar" aria-label="Cerrar">×</button></header>
    <div class="dlg-cuerpo">${cuerpo}</div>`);
}

function vistaCargando() {
  const e = Sync.estado;
  const error = e.red === 'error' || e.red === 'offline';
  return `<section class="centro">
    <div class="cargando" aria-hidden="true"></div>
    <h1>${error ? 'No se pudieron traer los datos' : 'Cargando datos de Google Sheets…'}</h1>
    <p class="sub">${error ? (e.red === 'offline' ? 'Este dispositivo no tiene internet. La primera carga necesita conexión.' : esc(e.error)) : 'La primera vez en este dispositivo tarda unos segundos.'}</p>
    ${error ? '<button class="btn" data-act="sync-ahora">Reintentar</button>' : ''}
  </section>`;
}

function vistaBienvenida() {
  return `<section class="centro">
    <h1>La hoja de Google no tiene facturas</h1>
    <p class="sub">${S.esAdmin() ? 'Puede cargar el Excel de cuentas por pagar desde Configuración.' : 'Un administrador debe cargar los datos. Cuando lo haga, aparecerán aquí solos.'}</p>
    ${S.esAdmin() ? '<a class="btn primario" href="#/config">Ir a Configuración</a>' : '<button class="btn" data-act="sync-ahora">Revisar ahora</button>'}
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
      <label class="buscar"><span class="vh">Buscar</span><input type="search" autocomplete="off" autocorrect="off" autocapitalize="off" spellcheck="false" enterkeyhint="search" placeholder="Buscar proveedor o número de factura" value="${esc(P.q)}" data-in="p-q"></label>
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
      <span class="apagar"><span class="estado prog">Programada ${money(prog.monto, f.moneda)} · ${fmtFecha(prog.fechaPago)}</span>
        ${f.revision ? `<button class="btn-revision on" data-act="p-revisar" data-v="${f.id}" title="${esc(f.observaciones || 'Pidió revisar')}">En revisión · quitar</button>`
          : `<button class="btn-revision" data-act="p-revisar" data-v="${f.id}">Pedir revisión</button>`}</span>
    </div>`;
  }
  return `<div class="fila-fact ${sel ? 'marcada' : ''}" data-fila="${f.id}">
    <label class="ck"><input type="checkbox" data-in="p-sel" data-v="${f.id}" ${sel ? 'checked' : ''}><span class="vh">Pagar factura ${esc(f.documento)}</span></label>
    <span class="doc">Fact. ${esc(f.documento)}${S.movimientosDe(f).length ? ` <span class="tag-abono" title="Ya tiene abonos">abonada</span>` : ''}
      ${f.revision
        ? `<button class="btn-revision on" data-act="p-revisar" data-v="${f.id}" title="${esc(f.observaciones || 'Pidió revisar')}">En revisión · quitar</button>`
        : `<button class="btn-revision" data-act="p-revisar" data-v="${f.id}">Pedir revisión</button>`}</span>
    <span class="venc">${f.fechaVence ? fmtFecha(f.fechaVence) : ''} ${etiquetaDias(f)}</span>
    <span class="saldo">${money(s, f.moneda)}</span>
    <span class="apagar">${sel ? `<label><span class="vh">Monto a pagar de la factura ${esc(f.documento)}</span>
      <span class="input-moneda"><span>${MONEDAS[f.moneda].simbolo}</span><input type="text" inputmode="decimal" data-monto value="${esc(ui.prog.sel.get(f.id))}" data-in="p-monto" data-v="${f.id}"></span></label>
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
    const m = parseMonto(txt), s = S.saldo(f);
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

const porRevisar = () => S.conSaldo().filter((f) => f.revision);

function seccionRevisar() {
  const lista = porRevisar().sort((a, b) => (a.fechaVence || '').localeCompare(b.fechaVence || ''));
  if (!lista.length) return '';
  return `<section class="lote revisar">
    <header class="lote-cab">
      <div><span class="eyebrow">Pedido por Norman</span><h2>Facturas para revisar</h2>
        <span class="sub">${lista.length} ${lista.length === 1 ? 'factura' : 'facturas'} · verifique que el registro esté bien y que no se haya pagado antes. Si todo está bien, márquela como revisada y pasa a pagarse aquí en «Por aplicar».</span></div>
    </header>
    <div class="tabla-scroll"><table class="tabla">
      <thead><tr><th>Sociedad</th><th>Proveedor</th><th>Factura</th><th class="num">Saldo</th><th class="num">A pagar</th><th>Qué revisar</th><th></th></tr></thead>
      <tbody>${lista.map((f) => {
        const p = S.programacionDe(f);
        return `<tr>
        <td>${esc(S.nombreSociedad(f.sociedad))}</td><td>${esc(S.proveedor(f.codigo).nombre)}</td>
        <td><a href="#/aplicar" data-act="f-abrir" data-v="${f.id}">${esc(f.documento)}</a></td>
        <td class="num">${money(S.saldo(f), f.moneda)}</td>
        <td class="num">${p ? `${money(p.monto, f.moneda)}<div class="sub">${fmtFecha(p.fechaPago)}</div>` : '<span class="sub">sin programar</span>'}</td>
        <td>${esc(f.observaciones) || '<span class="sub">—</span>'}</td>
        <td class="acciones-rev"><button class="btn chico primario" data-act="revisado" data-v="${f.id}">✓ Revisado</button>
          ${p ? `<button class="btn chico" data-act="rev-devolver" data-v="${f.id}">Devolver a Norman</button>` : ''}</td></tr>`;
      }).join('')}</tbody>
    </table></div>
  </section>`;
}

function dialogoPedirRevision(id) {
  const f = S.factura(id);
  const prog = S.programacionDe(f);
  const s = S.saldo(f);
  const marcado = ui.prog.sel.get(f.id);
  abrirDialogo(`
    <header class="dlg-cab"><div><h2>Pedir revisión · Factura ${esc(f.documento)}</h2>
      <p class="sub">${esc(S.proveedor(f.codigo).nombre)} · ${esc(S.nombreSociedad(f.sociedad))} · saldo ${money(s, f.moneda)}</p></div>
      <button class="btn-cerrar" data-act="cerrar" aria-label="Cerrar">×</button></header>
    <form class="dlg-cuerpo form-grid" data-form="pedir-revision" data-v="${f.id}">
      ${prog ? `<p class="ancho sub">Ya está programada por ${money(prog.monto, f.moneda)} para el ${fmtFecha(prog.fechaPago)}. No se paga hasta que la revisen.</p>`
        : `<label>Monto a pagar <input type="text" name="monto" inputmode="decimal" data-monto value="${esc(marcado || fmtNum(s))}" required></label>
      <label>Fecha de pago <input type="date" name="fecha" value="${ui.prog.fecha || hoyISO()}" required></label>`}
      <label class="ancho">¿Qué hay que revisar? <span class="sub-inline">opcional</span> <input name="nota" value="${esc(f.observaciones)}" placeholder="Ej. ¿ya se pagó? confirmar el monto"></label>
      <p class="ancho sub">Jonathan, Fernando y Daniel la verán en «Por aplicar». Cuando la marquen como revisada, queda lista para pagar.</p>
      <div class="ancho"><button class="btn primario" type="submit">Pedir revisión</button></div>
    </form>`);
}

function dialogoDevolverRevision(id) {
  const f = S.factura(id);
  abrirDialogo(`
    <header class="dlg-cab"><div><h2>Devolver a Norman · Factura ${esc(f.documento)}</h2>
      <p class="sub">${esc(S.proveedor(f.codigo).nombre)} · ${esc(S.nombreSociedad(f.sociedad))}</p></div>
      <button class="btn-cerrar" data-act="cerrar" aria-label="Cerrar">×</button></header>
    <form class="dlg-cuerpo form-grid" data-form="devolver-revision" data-v="${f.id}">
      <label class="ancho">¿Qué encontró? <input name="motivo" required placeholder="Ej. ya estaba pagada, el monto no coincide"></label>
      <p class="ancho sub">El pago sale de «Por aplicar» y no se paga. Norman ve el motivo en las observaciones de la factura.</p>
      <div class="ancho"><button class="btn primario" type="submit">Devolver</button></div>
    </form>`);
}

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
  ${seccionRevisar()}
  ${lotes.length ? lotes.map(tarjetaLote).join('') : `<p class="vacio grande">No hay pagos programados pendientes. Cuando Norman programe pagos, aparecerán aquí.</p>`}

  ${hechos.length ? `<section class="bloque aplicados">
    <h2>Aplicados recientemente</h2>
    <p class="sub">Toque un pago para adjuntar los comprobantes y compartir el aviso de pago con cada proveedor.</p>
    ${hechos.map(tarjetaAplicado).join('')}
  </section>` : ''}`;
}

/** Proveedores de un lote pendiente: los que ya se enviaron al banco y los que no. */
/**
 * Divide un lote: lo pendiente y cada archivo ya enviado al banco (un bloque por archivo).
 * Con `grupo`, `enviadas` son solo las líneas de ese archivo.
 */
function partesLote(l, grupo) {
  const pendientes = [], porGrupo = new Map();
  const parte = (ln, items, conExtra) => {
    const facturas = round2(items.reduce((t, x) => t + x.p.monto, 0)), extra = conExtra ? ln.extra : 0;
    return { ...ln, items, facturas, extra, total: round2(facturas + extra) };
  };
  for (const ln of l.lineas) {
    const resto = [], grupos = new Map();
    for (const it of ln.items) {
      const e = S.envioDe(it.p.id);
      if (!e) { resto.push(it); continue; }
      const g = String(e.grupo || 'x');
      if (!grupos.has(g)) grupos.set(g, { e, items: [] });
      grupos.get(g).items.push(it);
    }
    for (const [g, { e, items }] of grupos) {
      if (!porGrupo.has(g)) porGrupo.set(g, { grupo: g, e, lineas: [] });
      porGrupo.get(g).lineas.push(items.length === ln.items.length ? ln : parte(ln, items, false));
    }
    if (resto.length) pendientes.push(resto.length === ln.items.length ? ln : parte(ln, resto, true));
  }
  const envios = [...porGrupo.values()].sort((a, b) => String(a.e.fecha).localeCompare(String(b.e.fecha)));
  const elegidos = grupo ? envios.filter((x) => x.grupo === grupo) : envios;
  const enviadas = elegidos.flatMap((x) => x.lineas.map((ln) => ({ ln, e: x.e })));
  const forma = (ln) => ui.forma[l.clave + '|' + ln.codigo] || 'banco';
  return { envios, enviadas, pendientes, banco: pendientes.filter((ln) => forma(ln) === 'banco'), forma };
}

function tarjetaLote(l) {
  const { envios, pendientes, banco, forma } = partesLote(l);
  const nFact = l.lineas.reduce((s, x) => s + x.items.length, 0);
  const id = encodeURIComponent(l.clave);
  const subBanco = S.subLote(l, banco.map((x) => x.codigo));
  const problemas = banco.length ? S.problemasLote(subBanco) : [];
  const facts = (ln, conDevolver) => ln.items.map(({ p, f }) => `<div>${esc(f.documento)}: ${money(p.monto, f.moneda)}${p.monto < S.saldo(f) - 0.004 ? ' <span class="tag-abono">abono</span>' : ''}
      ${conDevolver ? `<button class="btn-link" data-act="a-devolver" data-v="${p.id}" title="Devolver a Norman">devolver</button>` : ''}</div>`).join('');
  return `<section class="lote">
    <header class="lote-cab">
      <div>
        <span class="eyebrow">Pagar el ${fmtFecha(l.fecha)}</span>
        <h2>${esc(S.nombreSociedad(l.sociedad))} · ${MONEDAS[l.moneda].nombre}</h2>
        <span class="sub">${l.lineas.length} ${l.lineas.length === 1 ? 'proveedor' : 'proveedores'} · ${nFact} facturas · debita ${l.cuentaDebito ? `<span class="mono">${fmtIBAN(l.cuentaDebito)}</span>` : '<em>cuenta sin configurar</em>'}</span>
      </div>
      <div class="lote-total">${money(l.total, l.moneda)}</div>
    </header>
    ${pendientes.length ? `
    <div class="tabla-scroll"><table class="tabla lote-tabla">
      <thead><tr><th>Proveedor</th><th>Facturas</th><th>IBAN</th><th class="num">Intereses / otros</th><th>Descripción para el banco</th><th class="num">Total</th><th>¿Cómo se paga?</th></tr></thead>
      <tbody>${pendientes.map((ln) => {
        const k = l.clave + '|' + ln.codigo, ke = encodeURIComponent(k), fm = forma(ln);
        return `<tr class="${fm === 'manual' ? 'fila-manual' : ''}">
        <td><strong>${esc(ln.prov.nombre)}</strong></td>
        <td class="facts">${facts(ln, true)}</td>
        <td class="mono">${ln.iban ? `${fmtIBAN(ln.iban)} ${errorIBAN(ln.iban) ? '<span class="malo">inválido</span>' : ''}` : `<a href="#/proveedores" data-act="ir-prov" data-v="${esc(ln.codigo)}" class="${fm === 'banco' ? 'malo' : ''}">Agregar IBAN</a>`}</td>
        <td class="num"><input class="mini" type="text" inputmode="decimal" data-monto value="${ln.extra ? fmtNum(ln.extra) : ''}" placeholder="0,00" data-in="a-extra" data-lote="${id}" data-v="${esc(ln.codigo)}" aria-label="Intereses u otros cargos para ${esc(ln.prov.nombre)}"></td>
        <td><input type="text" value="${esc(ln.descripcion)}" data-in="a-desc" data-lote="${id}" data-v="${esc(ln.codigo)}" aria-label="Descripción para ${esc(ln.prov.nombre)}"></td>
        <td class="num"><strong>${money(ln.total, l.moneda)}</strong></td>
        <td class="forma-celda">
          <span class="seg" role="group" aria-label="Cómo se paga ${esc(ln.prov.nombre)}">
            <button type="button" class="${fm === 'banco' ? 'on' : ''}" data-act="a-forma" data-v="${ke}|banco">Banco (.txt)</button><button type="button" class="${fm === 'manual' ? 'on' : ''}" data-act="a-forma" data-v="${ke}|manual">Manual</button></span>
          ${fm === 'manual' ? `<div class="manual-aplicar">
            <input type="text" id="mc-${ke}" placeholder="N° comprobante" aria-label="Comprobante del pago manual a ${esc(ln.prov.nombre)}" autocomplete="off">
            ${campoArchivo('man|' + k, 'PDF (opcional)')}
            <button class="btn primario chico" data-act="a-aplicar-manual" data-v="${ke}">Aplicar</button></div>` : ''}
        </td></tr>`;
      }).join('')}</tbody>
    </table></div>
    ${problemas.length ? `<div class="aviso"><strong>Antes de generar el archivo:</strong><ul>${problemas.map((p) => `<li>${esc(p)}</li>`).join('')}</ul></div>` : ''}
    <div class="lote-acciones">
      <div class="paso">
        <span class="num-paso">1</span>
        <span class="sub paso-txt">${banco.length ? `El archivo lleva <b>${banco.length} ${banco.length === 1 ? 'proveedor' : 'proveedores'}</b> · ${money(subBanco.total, l.moneda)}${banco.length < pendientes.length ? ' (sin los pagos manuales)' : ''}` : 'Todos los pagos pendientes de este grupo son manuales.'}</span>
        ${banco.length && !problemas.length ? `<label class="consec">Consecutivo del archivo <input type="text" inputmode="numeric" maxlength="3" id="consec-${id}" value="${S.siguienteConsecutivo(l.sociedad)}" aria-describedby="consec-ayuda-${id}"></label>
        <span class="sub consec-ayuda" id="consec-ayuda-${id}">Último usado en ${esc(S.nombreSociedad(l.sociedad))}: ${S.bancoSoc(l.sociedad).ultimo ? c3(S.bancoSoc(l.sociedad).ultimo) : '—'}</span>` : ''}
        <button class="btn" data-act="a-ver-txt" data-v="${id}" ${!banco.length || problemas.length ? 'disabled' : ''}>Ver contenido</button>
        <button class="btn oscuro" data-act="a-txt" data-v="${id}" ${!banco.length || problemas.length ? 'disabled' : ''}>Descargar archivo para el banco (.txt)</button>
      </div>
    </div>` : ''}
    ${envios.map(({ grupo, e, lineas }) => {
      const g = esc(grupo), k = `${id}-${g}`;
      return `
    <div class="enviado">
      <div class="enviado-cab"><span class="estado-envio">Enviado al banco · esperando comprobante</span>
        <span class="sub">${e.consecutivo ? `Consecutivo <b>${c3(e.consecutivo)}</b> · ` : ''}generado por ${esc(e.por || '—')} · ${fechaHora(Date.parse(e.fecha))} · ${lineas.length} ${lineas.length === 1 ? 'proveedor' : 'proveedores'} · ${money(round2(lineas.reduce((t, x) => t + x.total, 0)), l.moneda)}</span></div>
      <ul class="enviado-lista">${lineas.map((ln) => `<li><strong>${esc(ln.prov.nombre)}</strong> <span class="sub">${ln.items.map(({ f }) => esc(f.documento)).join(' · ')}</span> <span class="num">${money(ln.total, l.moneda)}</span></li>`).join('')}</ul>
      <div class="paso aplicar">
        <span class="num-paso">2</span>
        <label>Comprobante general del débito <input type="text" id="comp-${k}" placeholder="Del banco" autocomplete="off"></label>
        <label>Fecha <input type="date" id="fecha-${k}" value="${l.fecha}"></label>
        ${campoArchivo('gen|' + l.clave + '|' + grupo, 'PDF del banco (uso interno)')}
        <button class="btn primario" data-act="a-aplicar" data-v="${id}" data-g="${g}">Aplicar pago</button>
      </div>
      <div class="enviado-otras">
        <button class="btn-link" data-act="a-txt-rechazo" data-v="${id}" data-g="${g}">El banco rechazó el consecutivo: generar con otro</button> ·
        <button class="btn-link" data-act="a-txt-otra" data-v="${id}" data-g="${g}">Descargar el archivo otra vez</button> ·
        <button class="btn-link" data-act="a-volver" data-v="${id}" data-g="${g}">Volver a pendiente (si el banco lo rechazó)</button>
      </div>
    </div>`;
    }).join('')}
  </section>`;
}

/** Proveedores pagados en un lote aplicado, con sus facturas (desde los movimientos). */
function lineasAplicado(l) {
  const porProv = new Map();
  for (const m of Object.values(S.datos().movimientos)) {
    if (m.loteId !== l.id) continue;
    const f = S.factura(m.facturaId);
    if (!f) continue;
    if (!porProv.has(f.codigo)) porProv.set(f.codigo, { codigo: f.codigo, prov: S.proveedor(f.codigo), movs: [], total: 0 });
    const x = porProv.get(f.codigo);
    x.movs.push({ m, f });
    x.total = round2(x.total + m.monto);
  }
  return [...porProv.values()].sort((a, b) => a.prov.nombre.localeCompare(b.prov.nombre));
}

function tarjetaAplicado(l) {
  const abierto = !!ui.apAbierto[l.id];
  const gen = S.adjLote(l.id);
  const lineas = abierto ? lineasAplicado(l) : [];
  const docs = lineas.map((x) => S.docLinea(l.id, x.codigo));
  const conPdf = abierto ? docs.filter((d) => d.archivo?.id).length : 0;
  return `<div class="aplicado ${abierto ? 'abierto' : ''}">
    <button class="aplicado-cab" data-act="ap-toggle" data-v="${esc(l.id)}" aria-expanded="${abierto}">
      <span>${fmtFecha(l.fecha)}</span><span>${esc(S.nombreSociedad(l.sociedad))} · ${MONEDAS[l.moneda]?.nombre || ''}</span>
      <span class="mono">${esc(l.comprobante)}</span><span class="num">${money(l.total, l.moneda)}</span>
      <span class="sub">${nLineas(l)} ${nLineas(l) === 1 ? 'proveedor' : 'proveedores'} · ${esc(l.aplicadoPor)}</span><span aria-hidden="true">${abierto ? '▴' : '▾'}</span></button>
    ${abierto ? `<div class="aplicado-cuerpo">
      <div class="ap-general">📎 <b>Comprobante general del débito</b>
        ${gen?.id ? `<button class="btn-link" data-act="ver-arch" data-v="${esc(gen.id)}">Ver ${esc(gen.nombre)}</button>` : `<label class="adjuntar">Adjuntar PDF<input type="file" accept="application/pdf,image/jpeg,image/png" data-in="ap-gen" data-v="${esc(l.id)}" hidden></label>`}
        <span class="candado">🔒 Solo uso interno · no se comparte</span></div>
      <div class="tabla-scroll"><table class="tabla ap-tabla">
        <thead><tr><th>Proveedor</th><th>Facturas</th><th class="num">Monto</th><th>N° documento del banco</th><th>Comprobante del proveedor</th><th>Aviso de pago</th></tr></thead>
        <tbody>${lineas.map((x, i) => {
          const d = docs[i], v = esc(l.id + '|' + x.codigo);
          return `<tr><td><strong>${esc(x.prov.nombre)}</strong></td><td>${[...new Set(x.movs.map(({ f }) => f.documento))].map(esc).join(' · ')}</td>
          <td class="num">${money(x.total, l.moneda)}</td>
          <td><input type="text" class="mini-doc" value="${esc(d.documento || '')}" placeholder="opcional" data-in="ap-doc" data-v="${v}" aria-label="N° de documento del banco para ${esc(x.prov.nombre)}"></td>
          <td>${d.archivo?.id ? `<button class="btn-link" data-act="ver-arch" data-v="${esc(d.archivo.id)}">✓ ${esc(d.archivo.nombre)}</button>
              <label class="btn-link cambiar">cambiar<input type="file" accept="application/pdf,image/jpeg,image/png" data-in="ap-pdf" data-v="${v}" hidden></label>`
            : `<label class="adjuntar">📎 Adjuntar PDF<input type="file" accept="application/pdf,image/jpeg,image/png" data-in="ap-pdf" data-v="${v}" hidden></label>`}</td>
          <td><button class="btn primario chico" data-act="ap-aviso" data-v="${v}">${puedeCompartirArchivos() ? 'Compartir' : 'Descargar'}</button>
            ${d.archivo?.id ? '' : '<div class="sub mini-nota">sin comprobante: solo la página 1</div>'}</td></tr>`;
        }).join('')}</tbody>
      </table></div>
      <p class="sub">El aviso de pago lleva las facturas pagadas y, si lo adjuntó, el comprobante individual del proveedor. ${conPdf}/${lineas.length} con comprobante individual.</p>
    </div>` : ''}
  </div>`;
}

/** Aviso de pago (PDF) para un proveedor de un lote aplicado. */
async function avisoPago(loteId, codigo) {
  const l = S.datos().lotes[loteId];
  if (!l) throw new Error('No se encontró el pago');
  const x = lineasAplicado(l).find((y) => y.codigo === codigo);
  if (!x) throw new Error('No se encontró el proveedor en ese pago');
  const porFactura = new Map();
  let cargos = 0;
  for (const { m, f } of x.movs) {
    if (m.tipo === 'cargo') { cargos = round2(cargos + m.monto); continue; }
    if (!porFactura.has(f.id)) porFactura.set(f.id, { f, pagado: 0 });
    porFactura.get(f.id).pagado = round2(porFactura.get(f.id).pagado + m.monto);
  }
  const facturas = [...porFactura.values()].map(({ f, pagado }) => {
    // Saldo justo después de este pago (los pagos posteriores no cuentan)
    let saldo = f.monto, visto = false;
    for (const m of S.movimientosDe(f)) {
      if (m.tipo === 'cargo') continue;
      if (visto && m.loteId !== loteId) break;
      saldo = round2(saldo - m.monto);
      if (m.loteId === loteId) visto = true;
    }
    return { documento: f.documento, fechaFactura: fmtFecha(f.fechaFactura), monto: f.monto, pagado, saldo: Math.max(0, saldo), tipo: saldo <= 0.004 ? 'Pago total' : 'Abono' };
  }).sort((a, b) => String(a.documento).localeCompare(String(b.documento), 'es', { numeric: true }));
  const d = S.docLinea(loteId, codigo);
  let adjunto = null;
  if (d.archivo?.id) {
    const r = await Sync.leerArchivo(d.archivo.id);
    adjunto = { bytes: Uint8Array.from(atob(r.base64), (c) => c.charCodeAt(0)), tipo: r.tipo };
  }
  let desc = '';
  try { desc = (JSON.parse(l.detalle || '[]').find((y) => String(y.codigo) === String(codigo)) || {}).descripcion || ''; } catch { /* sin detalle */ }
  const iban = x.prov[l.moneda === 'USD' ? 'ibanUSD' : 'ibanCRC'];
  const ahora = new Date();
  const bytes = await crearAvisoPago({
    sociedad: nombreLegalSociedad(l.sociedad), proveedor: x.prov.nombre, fecha: fmtFecha(l.fecha), comprobante: l.comprobante,
    documento: d.documento || '', simbolo: l.moneda === 'USD' ? '$' : '¢', facturas, cargos, total: x.total,
    cuenta: iban ? fmtIBAN(iban) : '', descripcion: desc, fmtNum, adjunto,
    generado: `${fmtFecha(hoyISO())} ${String(ahora.getHours()).padStart(2, '0')}:${String(ahora.getMinutes()).padStart(2, '0')}`,
  });
  const nombre = `Aviso de pago - ${limpioArchivo(x.prov.nombre).slice(0, 60)} - ${limpioArchivo(l.comprobante)}.pdf`;
  return new File([bytes], nombre, { type: 'application/pdf' });
}

function loteDe(idCodificado) {
  const clave = decodeURIComponent(idCodificado);
  return S.lotesPendientes().find((l) => l.clave === clave);
}

function textoArchivo(l, consecutivo) {
  const soc = S.sociedades()[l.sociedad];
  return construirArchivo({
    cliente: S.bancoSoc(l.sociedad).cedula,
    consecutivo,
    moneda: l.moneda,
    fecha: l.fecha,
    control: S.cfg('controlEncabezado'),
    debito: { iban: l.cuentaDebito, descripcion: soc?.descripcionBanco || ('PROVEE ' + S.nombreSociedad(l.sociedad)).toUpperCase() },
    creditos: l.lineas.map((ln) => ({ iban: ln.iban, monto: ln.total, descripcion: ln.descripcion })),
  });
}

/** Descarga el .txt con ese consecutivo y lo deja anotado como el último usado de la sociedad. */
function descargarTxt(l, consecutivo, sufijo = '') {
  const txt = textoArchivo(l, consecutivo);
  const soc = S.nombreSociedad(l.sociedad).toUpperCase().replace(/[^A-Z0-9]+/g, '');
  descargar(`PROVEEDORES_${soc}_${l.moneda}_${l.fecha.split('-').reverse().join('')}_C${String(consecutivo).padStart(3, '0')}${sufijo}.txt`, txt, 'text/plain');
  if (Number(S.bancoSoc(l.sociedad).ultimo) !== Number(consecutivo)) S.guardarBancoSoc(l.sociedad, { ultimo: Number(consecutivo) });
}
const c3 = (n) => String(n).padStart(3, '0');
function leerConsecutivo(el) {
  const n = Number(String(el?.value || '').trim());
  if (!(Number.isInteger(n) && n >= 1 && n <= 999)) { avisar('El consecutivo del archivo debe ser un número de 1 a 999', 'error'); el?.focus(); return null; }
  return n;
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
    <div class="acciones">
      ${S.esNorman() ? '' : '<label class="btn">Importar facturas (CSV o Excel)<input type="file" accept=".csv,.txt,.xlsx,.xls" data-in="imp-archivo" hidden></label>'}
      <button class="btn primario" data-act="f-nueva">+ Nueva factura</button>
    </div>
  </div>
  <div class="filtros">
    <div class="fila-chips" role="group" aria-label="Sociedad">${chipsSociedad('f-soc', F.soc)}</div>
    <div class="fila-busqueda">
      <label class="buscar"><span class="vh">Buscar</span><input type="search" autocomplete="off" autocorrect="off" autocapitalize="off" spellcheck="false" enterkeyhint="search" placeholder="Proveedor, factura o comprobante" value="${esc(F.q)}" data-in="f-q"></label>
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

function ultimoPago(f) {
  return [...S.movimientosDe(f)].reverse().find((m) => m.tipo !== 'cargo') || null;
}

function vistaCanceladas() {
  const C = ui.canc;
  const texto = C.q.trim().toLowerCase();
  let lista = S.listaFacturas().filter((f) => S.saldo(f) <= 0.004 && f.monto > 0).filter((f) => {
    if (C.soc && f.sociedad !== C.soc) return false;
    if (C.mon && f.moneda !== C.mon) return false;
    if (texto) {
      const pagos = S.movimientosDe(f).map((m) => m.comprobante).join(' ');
      if (!(`${S.proveedor(f.codigo).nombre} ${f.documento} ${pagos}`.toLowerCase().includes(texto))) return false;
    }
    return true;
  }).map((f) => ({ f, pago: ultimoPago(f) }));
  lista.sort((a, b) => String(b.pago?.fecha || '').localeCompare(String(a.pago?.fecha || '')) || String(b.f.fechaFactura || '').localeCompare(String(a.f.fechaFactura || '')));
  const total = lista.length;
  const t = { CRC: 0, USD: 0 };
  for (const x of lista) t[x.f.moneda] = round2(t[x.f.moneda] + x.f.monto);
  lista = lista.slice(0, C.limite);
  return `
  <div class="titulo-vista"><div><h1>Facturas canceladas</h1><p class="sub">Todas las facturas pagadas por completo, incluidas las que venían en el Excel</p></div></div>
  <div class="filtros">
    <div class="fila-chips" role="group" aria-label="Sociedad">${chipsSociedad('c-soc', C.soc)}</div>
    <div class="fila-busqueda">
      <label class="buscar"><span class="vh">Buscar</span><input type="search" autocomplete="off" autocorrect="off" autocapitalize="off" spellcheck="false" enterkeyhint="search" placeholder="Proveedor, número de factura o comprobante" value="${esc(C.q)}" data-in="c-q"></label>
      <label>Moneda <select data-in="c-mon"><option value="">Ambas</option><option value="CRC" ${C.mon === 'CRC' ? 'selected' : ''}>Colones</option><option value="USD" ${C.mon === 'USD' ? 'selected' : ''}>Dólares</option></select></label>
    </div>
  </div>
  <p class="resumen-lista">${total} facturas canceladas · ${totalesTxt(t)}</p>
  <div class="tabla-scroll"><table class="tabla clic">
    <thead><tr><th>Sociedad</th><th>Proveedor</th><th>Factura</th><th>Fecha factura</th><th class="num">Monto</th><th>Pagada el</th><th>Comprobante</th></tr></thead>
    <tbody>${lista.map(({ f, pago }) => `<tr data-act="f-abrir" data-v="${f.id}" tabindex="0">
      <td>${esc(S.nombreSociedad(f.sociedad))}</td><td>${esc(S.proveedor(f.codigo).nombre)}</td>
      <td><a href="#/canceladas" data-act="f-abrir" data-v="${f.id}">${esc(f.documento)}</a></td>
      <td>${fmtFecha(f.fechaFactura)}</td><td class="num">${money(f.monto, f.moneda)}</td>
      <td>${pago?.fecha ? fmtFecha(pago.fecha) : '<span class="sub">sin fecha</span>'}</td><td>${esc(pago?.comprobante || '')}</td></tr>`).join('') || '<tr><td colspan="7" class="sub">No hay facturas canceladas con esos filtros.</td></tr>'}</tbody>
  </table></div>
  ${total > lista.length ? `<p class="centro-txt"><button class="btn" data-act="c-mas">Mostrar más (${total - lista.length} restantes)</button></p>` : ''}`;
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
      ${s > 0.004 && !S.esNorman() ? `<div class="acciones-fila"><button class="btn btn-norman" data-act="pn-abrir" data-v="${f.id}">✔ Pagar ahora</button></div>` : ''}
      <form class="form-grid" data-form="f-fechas" data-v="${f.id}">
        <label>Fecha de factura <input type="date" name="fechaFactura" value="${f.fechaFactura || ''}"></label>
        <label>Vencimiento ${f.venceEstimado ? '<span class="est">EST</span>' : ''}<input type="date" name="fechaVence" value="${f.fechaVence || ''}"></label>
        <label class="ancho">Observaciones <input type="text" name="observaciones" value="${esc(f.observaciones)}"></label>
${f.revision ? `        <p class="ancho aviso-pruebas">Norman pidió revisar esta factura. Se marca como revisada en «Por aplicar».</p>
` : ''}        <div class="ancho"><button class="btn" type="submit">Guardar cambios</button></div>
      </form>
      <h3>Movimientos</h3>
      ${S.movimientosDe(f).length ? `<div class="tabla-scroll"><table class="tabla">
        <thead><tr><th>Fecha</th><th>Tipo</th><th>Comprobante</th><th class="num">Monto</th><th class="num">Saldo después</th><th>Registró</th></tr></thead>
        <tbody>${S.movimientosDe(f).map((m) => {
          if (m.tipo !== 'cargo') acumulado = round2(acumulado - m.monto);
          const prog = m.loteId ? S.datos().programaciones[String(m.id).slice(String(m.loteId).length + 1)] : null;
          const tagNorman = S.esAutorizadoNorman(prog) ? ' <span class="tag-norman">Pagar ahora</span>' : '';
          const tipoTxt = String(prog?.nota || '').startsWith(S.NOTA_NC) ? 'Nota de crédito' : (tipos[m.tipo] || m.tipo);
          return `<tr><td>${fmtFecha(m.fecha) || '<span class="sub">sin fecha</span>'}</td><td>${tipoTxt}${tagNorman}</td><td>${esc(m.comprobante)}${linkComprobante(S.comprobanteDe(m))}</td>
            <td class="num">${money(m.monto, f.moneda)}</td><td class="num">${m.tipo === 'cargo' ? '' : money(acumulado, f.moneda)}</td><td>${esc(m.por)}</td></tr>`;
        }).join('')}</tbody></table></div>` : '<p class="sub">Sin movimientos.</p>'}
      ${s > 0.004 ? `
      <h3>Registrar abono o nota de crédito</h3>
      <form class="form-grid" data-form="f-mov" data-v="${f.id}">
        <label>Tipo <select name="tipo" data-in="fm-tipo"><option value="abono">Abono / pago</option><option value="nc">Nota de crédito</option></select></label>
        <label>Monto <input type="text" name="monto" inputmode="decimal" data-monto value="${fmtNum(s)}" data-in="fm-monto" data-v="${f.id}">
          <span class="pn-resto" id="fm-resto">${textoResto(f, s)}</span></label>
        <label>Fecha <input type="date" name="fecha" value="${hoyISO()}"></label>
        <label>N° comprobante o de nota <input type="text" name="comprobante" required autocomplete="off"></label>
        <div class="ancho"><button class="btn primario" type="submit">Registrar</button></div>
      </form>` : ''}
    </div>`);
}

// ---------------------------------------------------------------- Pago autorizado por Norman

function dialogoPagoNorman(id) {
  const f = S.factura(id);
  const prov = S.proveedor(f.codigo);
  const s = S.saldo(f);
  const ya = S.programacionDe(f);
  const cab = `<header class="dlg-cab">
      <div><h2>Pagar ahora</h2>
        <p class="sub">Factura ${esc(f.documento)} · ${esc(prov.nombre)} · ${esc(S.nombreSociedad(f.sociedad))} · saldo ${money(s, f.moneda)}</p></div>
      <button class="btn-cerrar" data-act="cerrar" aria-label="Cerrar">×</button></header>`;
  if (ya) {
    abrirDialogo(`${cab}<div class="dlg-cuerpo">
      <p class="aviso-pruebas">Esta factura ya está programada por ${esc(ya.por)} por ${money(ya.monto, f.moneda)} para el ${fmtFecha(ya.fechaPago)}${S.envioDe(ya.id) ? ' y ya se envió al banco' : ''}. Aplíquela desde «Por aplicar» para no pagarla dos veces.</p>
      <div class="acciones-fila"><a class="btn primario" href="#/aplicar" data-act="cerrar">Ir a Por aplicar</a></div></div>`);
    return;
  }
  const lote = S.loteDirecto(f, s, hoyISO());
  const prob = S.problemasLote(lote);
  abrirDialogo(`${cab}
    <form class="dlg-cuerpo form-grid pagar-ahora" data-form="pago-norman" data-v="${f.id}" data-forma="banco">
      <div class="ancho"><span class="lbl">¿Cómo se paga?</span>
        <div class="pa-formas">
          <label class="pa-forma"><input type="radio" name="forma" value="banco" checked data-in="pn-forma"><span><b>Banco (.txt)</b>Genera el archivo para el banco. El comprobante se escribe en «Por aplicar» cuando el banco aplique.</span></label>
          <label class="pa-forma"><input type="radio" name="forma" value="manual" data-in="pn-forma"><span><b>Manual</b>Transferencia hecha a mano (por ejemplo, cuando Norman lo pide). El comprobante se escribe ya.</span></label>
        </div></div>
      <label>Monto a pagar <input type="text" name="monto" inputmode="decimal" data-monto value="${fmtNum(s)}" required data-in="pn-monto" data-v="${f.id}">
        <span class="pn-resto" id="pn-resto">${textoResto(f, s)}</span></label>
      <label>Fecha de pago <input type="date" name="fecha" value="${hoyISO()}" required></label>
      <label class="ancho">Nota <span class="sub-inline">opcional</span> <input type="text" name="nota" placeholder="Ej. Norman lo pidió por teléfono"></label>
      <div class="ancho solo-banco pn-banco">
        <label>Descripción en el banco <input type="text" name="descripcion" maxlength="60" value="${esc(lote.lineas[0].descripcion)}"></label>
        <label>Consecutivo del archivo <input type="text" name="consecutivo" inputmode="numeric" maxlength="3" value="${S.siguienteConsecutivo(f.sociedad)}">
          <span class="sub">Último usado en ${esc(S.nombreSociedad(f.sociedad))}: ${S.bancoSoc(f.sociedad).ultimo ? c3(S.bancoSoc(f.sociedad).ultimo) : '—'}</span></label>
        ${prob.length ? `<div class="aviso"><strong>Para generar el archivo falta:</strong><ul>${prob.map((x) => `<li>${esc(x)}</li>`).join('')}</ul></div>` : ''}
        <p class="sub">Al tocar el botón se descarga el .txt y el pago queda «Enviado al banco · esperando comprobante» en Por aplicar.</p>
      </div>
      <label class="ancho solo-manual">N° de comprobante <input type="text" name="comprobante" autocomplete="off"></label>
      <div class="ancho solo-manual">${campoArchivo('pn|' + f.id, 'Comprobante del banco (PDF o foto) · opcional')}</div>
      <div class="ancho acciones-fila">
        <button class="btn oscuro solo-banco" type="submit" ${prob.length ? 'disabled' : ''}>Descargar .txt y enviar al banco</button>
        <button class="btn primario solo-manual" type="submit">Aplicar pago</button>
        <button class="btn" type="button" data-act="f-abrir" data-v="${f.id}">Volver</button></div>
    </form>`);
}

/** Casilla para elegir un comprobante (PDF o foto). El archivo se guarda en memoria hasta aplicar. */
function campoArchivo(clave, texto) {
  const a = ui.adj[clave];
  return `<label class="adjuntar ${a ? 'listo' : ''}">📎 <span>${a ? `<b>${esc(a.name)}</b> · ${Math.max(1, Math.round(a.size / 1024))} KB` : esc(texto)}</span>
    <input type="file" accept="application/pdf,image/jpeg,image/png" data-in="adj-file" data-k="${esc(clave)}" hidden></label>`;
}

async function archivoBase64(file) {
  if (file.size > 8 * 1024 * 1024) throw new Error('El archivo pesa más de 8 MB');
  return new Promise((ok, mal) => {
    const fr = new FileReader();
    fr.onload = () => ok(String(fr.result).split(',')[1]);
    fr.onerror = () => mal(new Error('No se pudo leer el archivo'));
    fr.readAsDataURL(file);
  });
}

/** Sube un comprobante a Drive y devuelve { id, nombre, tipo }. */
async function subirComprobante(file, nombre) {
  const base64 = await archivoBase64(file);
  const ext = file.type === 'image/png' ? '.png' : file.type === 'image/jpeg' ? '.jpg' : '.pdf';
  const r = await Sync.subirArchivo({ nombre: (nombre || file.name.replace(/\.[^.]+$/, '')) + ext, tipo: file.type || 'application/pdf', base64 });
  return { id: r.id, nombre: r.nombre, tipo: r.tipo };
}

/** Sube el archivo elegido (si hay) y lo anota. No detiene el pago si falla: se puede adjuntar después. */
async function adjuntarDespues(clave, nombre, anotar) {
  const file = ui.adj[clave];
  if (!file) return;
  delete ui.adj[clave];
  avisar('Subiendo comprobante…');
  try { const a = await subirComprobante(file, nombre); const r = anotar(a); if (r?.error) throw new Error(r.error); avisar('Comprobante adjuntado'); render(); }
  catch (err) { avisar('El pago quedó aplicado, pero no se pudo adjuntar el comprobante: ' + err.message + '. Puede adjuntarlo en «Aplicados recientemente».', 'error'); }
}

/** Muestra un comprobante guardado (se trae del servicio; no queda guardado en el dispositivo). */
async function verComprobante(id) {
  abrirDialogo(`<header class="dlg-cab"><h2>Comprobante</h2><button class="btn-cerrar" data-act="cerrar" aria-label="Cerrar">×</button></header>
    <div class="dlg-cuerpo"><p class="vacio">Abriendo comprobante…</p></div>`);
  try {
    const r = await Sync.leerArchivo(id);
    const bytes = Uint8Array.from(atob(r.base64), (c) => c.charCodeAt(0));
    mostrarArchivo(new File([bytes], r.nombre, { type: r.tipo }), 'Comprobante', r.nombre);
  } catch (err) { cerrarDialogo(); avisar(err.message, 'error'); }
}

/** Enlace «📎 Ver comprobante» (no se imprime). */
function linkComprobante(a) {
  if (!a?.id) return '';
  return ` <button class="btn-link ver-arch no-imprimir" data-act="ver-arch" data-v="${esc(a.id)}" title="${a.individual ? 'Comprobante del proveedor' : 'Comprobante general del débito (uso interno)'}">📎 Ver comprobante</button>`;
}

/** Vista previa de un archivo (PDF o imagen) con Compartir y Descargar. */
function mostrarArchivo(archivo, titulo, sub) {
  if (ui.urlArchivo) URL.revokeObjectURL(ui.urlArchivo);
  const url = ui.urlArchivo = URL.createObjectURL(archivo);
  ui.archivoAbierto = archivo;
  abrirDialogo(`<header class="dlg-cab"><div><h2>${esc(titulo)}</h2><p class="sub">${esc(sub || archivo.name)}</p></div><button class="btn-cerrar" data-act="cerrar" aria-label="Cerrar">×</button></header>
    <div class="dlg-cuerpo ver-comprobante">
      ${archivo.type.startsWith('image/') ? `<img src="${url}" alt="${esc(titulo)}">` : `<iframe src="${url}" title="${esc(titulo)}"></iframe>`}
      <div class="acciones-fila">
        ${puedeCompartirArchivos() ? '<button class="btn primario" data-act="arch-compartir">Compartir (WhatsApp…)</button>' : ''}
        <button class="btn${puedeCompartirArchivos() ? '' : ' primario'}" data-act="arch-descargar">Descargar</button>
        <a class="btn" href="${url}" target="_blank" rel="noopener">Abrir en otra pestaña</a></div>
    </div>`);
}

/** Debajo del monto: si es abono, cuánto queda; si es el saldo completo, que queda cancelada. */
function textoResto(f, monto, nc = false) {
  const s = S.saldo(f);
  if (!(monto > 0)) return '<span class="pn-mal">Escriba el monto.</span>';
  if (monto > s + 0.004) return `<span class="pn-mal">Es mayor que el saldo (${money(s, f.moneda)}).</span>`;
  if (monto >= s - 0.004) return `<span class="pn-ok">${nc ? 'Nota de crédito por el total' : 'Pago total'}: la factura queda cancelada.</span>`;
  return `<span class="pn-abono">${nc ? 'Nota de crédito' : 'Abono'} de ${money(monto, f.moneda)} · queda un saldo de <strong>${money(round2(s - monto), f.moneda)}</strong></span>`;
}
function actualizarRestoMov() {
  const form = document.querySelector('[data-form="f-mov"]');
  const r = document.getElementById('fm-resto');
  if (!form || !r) return;
  r.innerHTML = textoResto(S.factura(form.dataset.v), parseMonto(form.monto.value), form.tipo.value === 'nc');
}

function valoresPagoNorman(form) {
  const d = Object.fromEntries(new FormData(form));
  return { forma: d.forma || 'banco', monto: parseMonto(d.monto), fecha: d.fecha, nota: d.nota, comprobante: d.comprobante, descripcion: d.descripcion };
}

function seccionAutorizados() {
  const lista = S.autorizadosNorman(30);
  if (!lista.length) return '';
  return `<section class="bloque bloque-norman">
    <h2>Pagados con «Pagar ahora» <span class="sub-inline">últimos 30 días</span></h2>
    <p class="sub">Pagos hechos desde la factura, por banco o manual (por ejemplo, cuando Norman lo indica de palabra).</p>
    <div class="tabla-scroll"><table class="tabla">
      <thead><tr><th>Fecha</th><th>Sociedad</th><th>Proveedor</th><th>Factura</th><th class="num">Monto</th><th>Aplicó</th><th>Comprobante</th><th>Nota</th></tr></thead>
      <tbody>${lista.map(({ p, f }) => `<tr data-act="f-abrir" data-v="${f.id}" tabindex="0">
        <td>${fmtFecha(p.fechaPago)}</td><td>${esc(S.nombreSociedad(f.sociedad))}</td><td>${esc(S.proveedor(f.codigo).nombre)}</td>
        <td>${esc(f.documento)}</td><td class="num">${money(p.monto, f.moneda)}</td><td>${esc(p.cerradoPor)}</td>
        <td>${esc(p.comprobante)}</td><td>${esc(p.nota.replace(/^(Autorizado por Norman|Pagar ahora)\s*·?\s*/, ''))}</td></tr>`).join('')}</tbody>
    </table></div>
  </section>`;
}

// ---------------------------------------------------------------- Importar facturas

const letraCol = (i) => { let s = ''; i += 1; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
const firmaEncabezado = (enc) => 'cxp-mapeo:' + enc.map((h) => String(h).trim().toLowerCase()).join('|');

async function cargarArchivoImportar(file) {
  const esTexto = /\.(csv|txt)$/i.test(file.name);
  const contenido = esTexto ? await leerTexto(file) : await file.arrayBuffer();
  const leidas = leerFilas(contenido, window.XLSX, esTexto);
  if (!leidas.length) throw new Error('El archivo no tiene filas de facturas');
  const gti = detectarGTI(leidas);
  const filas = gti ? gti.filas : leidas;
  const inicio = gti ? gti.inicio : filaEncabezado(filas);
  const enc = filas[inicio].map((h) => String(h).trim());
  let guardado = null;
  try { guardado = JSON.parse(localStorage.getItem(firmaEncabezado(enc)) || 'null'); } catch { /* sin almacenamiento */ }
  ui.imp = {
    nombre: file.name, filas, inicio, enc, gti: !!gti, asignados: {},
    mapeo: gti ? gti.mapeo : (guardado?.mapeo || sugerirMapeo(enc)), fijos: guardado?.fijos || { sociedad: '', moneda: '' },
  };
  dialogoImportar();
}

/** Lee un CSV respetando tildes: prueba UTF-8 y, si hay caracteres dañados, Windows-1252 (Excel). */
async function leerTexto(file) {
  const buf = await file.arrayBuffer();
  const utf8 = new TextDecoder('utf-8').decode(buf);
  return utf8.includes('\uFFFD') ? new TextDecoder('windows-1252').decode(buf) : utf8;
}

function dialogoImportar() {
  const I = ui.imp;
  const d = S.datos();
  const res = prepararFacturas(I.filas, I.inicio, I.mapeo, I.fijos, d, { asignados: I.asignados, memoria: memoriaCedulas() });
  I.res = res;
  const opcionesCol = (k, opcional) => `<option value="">${opcional ? '— no viene en el archivo —' : '— elegir columna —'}</option>` +
    I.enc.map((h, i) => `<option value="${i}" ${String(I.mapeo[k]) === String(i) ? 'selected' : ''}>${letraCol(i)} · ${esc(h || '(sin título)')}</option>`).join('');
  const socs = Object.keys(S.sociedades());
  const campo = (c) => {
    const sinCol = I.mapeo[c.k] == null || I.mapeo[c.k] === '';
    let extra = '';
    if (c.k === 'sociedad' && sinCol) extra = `<select data-in="imp-fijo" data-v="sociedad" aria-label="Sociedad para todas las facturas"><option value="">— elegir sociedad —</option>${socs.map((x) => `<option value="${esc(x)}" ${I.fijos.sociedad === x ? 'selected' : ''}>${esc(S.nombreSociedad(x))}</option>`).join('')}</select>`;
    if (c.k === 'moneda' && sinCol) extra = `<select data-in="imp-fijo" data-v="moneda" aria-label="Moneda para todas las facturas"><option value="">— elegir moneda —</option><option value="CRC" ${I.fijos.moneda === 'CRC' ? 'selected' : ''}>Colones</option><option value="USD" ${I.fijos.moneda === 'USD' ? 'selected' : ''}>Dólares</option></select>`;
    return `<label>${esc(c.n)}${c.opcional ? ' <span class="sub-inline">opcional</span>' : ''}
      <select data-in="imp-col" data-v="${c.k}">${opcionesCol(c.k, c.opcional || c.fijo)}</select>
      ${extra ? `<span class="sub">${c.k === 'sociedad' ? 'Para todas las facturas del archivo:' : 'Para todas las facturas del archivo:'}</span>${extra}` : ''}</label>`;
  };
  const r = res.resumen;
  const muestra = res.filas.slice(0, 60);
  const etiqueta = { nueva: '<span class="estado canc">Nueva</span>', existe: '<span class="estado pend">Ya existe</span>', omitida: '<span class="estado">No se agrega</span>', error: '<span class="estado rev">Revisar</span>' };
  abrirDialogo(`
    <header class="dlg-cab"><div><h2>Importar facturas</h2><p class="sub">${esc(I.nombre)} · ${I.filas.length - I.inicio - 1} filas</p></div>
      <button class="btn-cerrar" data-act="cerrar" aria-label="Cerrar">×</button></header>
    <div class="dlg-cuerpo">
      ${I.gti ? `<p class="imp-gti">✓ Reporte de GTI reconocido. Se toma el <strong>TotalComprobante</strong>, la moneda, la fecha de emisión y el número del consecutivo (por ejemplo, 00100001010000107091 → 107091).</p>
      <details class="imp-detalles"><summary>Ver columnas</summary>` : '<h3>1 · ¿Qué columna es cada dato?</h3>'}
      <p class="sub">La aplicación propone las columnas por su título. Corrija las que no correspondan; se recordarán para la próxima vez.</p>
      <div class="form-grid imp-mapeo">${CAMPOS_IMP.map(campo).join('')}</div>
      ${I.gti ? '</details>' : ''}

      ${seccionEmisores(res.emisores)}

      <h3>${I.gti ? '' : '2 · '}Revise las facturas</h3>
      <div class="imp-resumen">
        <span class="estado canc">${r.nuevas} nuevas</span>
        ${r.existentes ? `<span class="estado pend">${r.existentes} ya existen</span>` : ''}
        ${r.omitidas ? `<span class="estado">${r.omitidas} notas de crédito (no se agregan)</span>` : ''}
        ${r.errores ? `<span class="estado rev">${r.errores} con datos por revisar</span>` : ''}
        ${r.proveedoresNuevos.length ? `<span class="estado prog">${r.proveedoresNuevos.length} proveedores nuevos</span>` : ''}
      </div>
      <div class="tabla-scroll"><table class="tabla">
        <thead><tr><th>Fila</th><th>Estado</th><th>Sociedad</th><th>Proveedor</th><th>Factura</th><th>Fecha</th><th>Vence</th><th class="num">Monto</th></tr></thead>
        <tbody>${muestra.map((x) => `<tr>
          <td>${x.n}</td><td>${etiqueta[x.estado]}${x.mensaje ? `<div class="sub">${esc(x.mensaje)}</div>` : ''}</td>
          <td>${esc(x.factura.sociedad ? S.nombreSociedad(x.factura.sociedad) : '')}</td>
          <td>${esc(x.proveedor?.nombre || x.nombreProv || '')}${x.proveedorNuevo ? ' <span class="tag-abono">nuevo</span>' : ''}</td>
          <td>${esc(x.factura.documento)}</td><td>${fmtFecha(x.factura.fechaFactura)}</td><td>${fmtFecha(x.factura.fechaVence)}</td>
          <td class="num">${x.factura.moneda ? money(x.factura.monto, x.factura.moneda) : fmtNum(x.factura.monto || 0)}</td></tr>`).join('')}</tbody>
      </table></div>
      ${res.filas.length > muestra.length ? `<p class="sub">Se muestran las primeras ${muestra.length} de ${res.filas.length} filas.</p>` : ''}

      <div class="acciones-fila imp-acciones">
        <button class="btn primario" data-act="imp-confirmar" ${r.nuevas ? '' : 'disabled'}>Agregar ${r.nuevas} ${r.nuevas === 1 ? 'factura' : 'facturas'}</button>
        <button class="btn" data-act="cerrar">Cancelar</button>
      </div>
      <p class="sub">Solo se agregan las filas «Nueva». Las que ya existen, las notas de crédito y las que hay que revisar se quedan fuera.</p>
    </div>`);
}

/** Cédula → código de proveedor elegido antes en este dispositivo. */
function memoriaCedulas() {
  try { return JSON.parse(localStorage.getItem('cxp-cedulas') || '{}'); } catch { return {}; }
}

function seccionEmisores(emisores) {
  const lista = emisores.filter((e) => e.filas > 0 || e.clave in ui.imp.asignados);
  if (!lista.length) return '';
  const provs = Object.values(S.datos().proveedores).sort((a, b) => a.nombre.localeCompare(b.nombre));
  const como = { cedula: 'reconocido por la cédula', nombre: 'reconocido por el nombre', codigo: 'por el código', elegido: 'elegido por usted', nuevo: 'no se encontró: se creará' };
  return `<h3>Proveedores del archivo</h3>
    <p class="sub">El archivo trae la cédula, no su código contable. Confirme cada proveedor o elija el correcto; la aplicación lo recordará.</p>
    <div class="imp-emisores">${lista.map((e) => `
      <div class="imp-emisor ${e.codigo ? '' : 'es-nuevo'}">
        <div><strong>${esc(e.nombre)}</strong><span class="sub">${e.cedula ? `Céd. ${esc(e.cedula)} · ` : ''}${e.filas} ${e.filas === 1 ? 'factura' : 'facturas'} · ${como[e.como]}</span></div>
        <select data-in="imp-prov" data-v="${esc(e.clave)}" aria-label="Proveedor para ${esc(e.nombre)}">
          <option value="" ${e.codigo ? '' : 'selected'}>➕ Crear proveedor nuevo${e.nuevo ? ` (código ${esc(e.nuevo.codigo)})` : ''}</option>
          ${provs.map((p) => `<option value="${esc(p.codigo)}" ${p.codigo === e.codigo ? 'selected' : ''}>${esc(p.codigo)} · ${esc(p.nombre)}</option>`).join('')}
        </select>
      </div>`).join('')}</div>`;
}

function confirmarImportacion() {
  const I = ui.imp;
  const nuevas = I.res.filas.filter((x) => x.estado === 'nueva');
  if (!nuevas.length) return;
  try {
    if (!I.gti) localStorage.setItem(firmaEncabezado(I.enc), JSON.stringify({ mapeo: I.mapeo, fijos: I.fijos }));
    const mem = memoriaCedulas();
    for (const e of I.res.emisores) if (e.cedula && e.codigo) mem[e.cedula] = e.codigo;
    localStorage.setItem('cxp-cedulas', JSON.stringify(mem));
  } catch { /* sin almacenamiento */ }
  const errores = [];
  for (const p of I.res.resumen.proveedoresNuevos) {
    const r = S.guardarProveedor(p.codigo, { nombre: p.nombre, alias: '', ibanCRC: '', ibanUSD: '', plazo: 30 }, true);
    if (r.error) errores.push(p.nombre + ': ' + r.error);
  }
  let ok = 0;
  for (const x of nuevas) {
    const r = S.nuevaFactura(x.factura);
    if (r.error) errores.push(`Fila ${x.n}: ${r.error}`); else ok++;
  }
  cerrarDialogo();
  ui.imp = null;
  avisar(`Se agregaron ${ok} facturas` + (errores.length ? ` · ${errores.length} con problemas: ${errores.slice(0, 3).join(' · ')}` : ''), errores.length ? 'error' : 'ok');
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
      <label>Monto <input type="text" name="monto" inputmode="decimal" data-monto required autocomplete="off"></label>
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
      <label class="buscar"><span class="vh">Buscar</span><input type="search" autocomplete="off" autocorrect="off" autocapitalize="off" spellcheck="false" enterkeyhint="search" placeholder="Nombre, alias o código" value="${esc(P.q)}" data-in="pv-q"></label>
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

// Color de cada sociedad (en el orden en que aparecen) y de cada tramo de antigüedad.
const COLORES_SOC = ['soc-1', 'soc-2', 'soc-3', 'soc-4', 'soc-5', 'soc-6'];
const colorSociedad = (s) => COLORES_SOC[Object.keys(S.sociedades()).indexOf(s) % COLORES_SOC.length] || 'soc-1';

function barra(valor, max, clase) {
  const pct = max > 0 ? Math.max(valor > 0 ? 2 : 0, Math.round((valor / max) * 100)) : 0;
  return `<span class="barra ${clase}" aria-hidden="true"><span style="width:${pct}%"></span></span>`;
}

function vistaTablero() {
  const pend = S.conSaldo();
  const venc = pend.filter((f) => S.diasVencida(f) > 0);
  const semana = pend.filter((f) => { const d = S.diasVencida(f); return d != null && d <= 0 && d > -7; });
  const tramos = [['Por vencer', (d) => d <= 0, 't-0'], ['1–30 días', (d) => d >= 1 && d <= 30, 't-1'], ['31–60 días', (d) => d > 30 && d <= 60, 't-2'], ['61–90 días', (d) => d > 60 && d <= 90, 't-3'], ['Más de 90', (d) => d > 90, 't-4']];
  const socs = Object.keys(S.sociedades());
  const porSoc = socs.map((s) => ({ s, t: S.totales(pend.filter((f) => f.sociedad === s)) }));
  const maxCRC = Math.max(0, ...porSoc.map((x) => x.t.CRC)), maxUSD = Math.max(0, ...porSoc.map((x) => x.t.USD));
  const porTramo = tramos.map(([n, fn, c]) => ({ n, c, t: S.totales(pend.filter((f) => { const d = S.diasVencida(f); return d != null && fn(d); })) }));
  const maxTC = Math.max(0, ...porTramo.map((x) => x.t.CRC)), maxTU = Math.max(0, ...porTramo.map((x) => x.t.USD));
  return `
  <div class="titulo-vista"><div><h1>Tablero</h1><p class="sub">Al ${fmtFecha(hoyISO())}</p></div></div>
  <div class="kpis">
    <div class="kpi k-total"><span>Total por pagar</span><strong>${money(S.totales(pend).CRC, 'CRC')}</strong><strong class="usd">${money(S.totales(pend).USD, 'USD')}</strong><small>${pend.length} facturas</small></div>
    <div class="kpi alerta"><span>Vencido</span><strong>${money(S.totales(venc).CRC, 'CRC')}</strong><strong class="usd">${money(S.totales(venc).USD, 'USD')}</strong><small>${venc.length} facturas</small></div>
    <div class="kpi k-semana"><span>Vence en 7 días</span><strong>${money(S.totales(semana).CRC, 'CRC')}</strong><strong class="usd">${money(S.totales(semana).USD, 'USD')}</strong><small>${semana.length} facturas</small></div>
  </div>
  <div class="dos-col">
    <section class="bloque"><h2>Saldo por sociedad</h2>
      <div class="tarjetas-soc">${porSoc.map(({ s, t }) => `
        <a class="tarjeta-soc ${colorSociedad(s)}" href="#/facturas" data-act="ir-sociedad" data-v="${esc(s)}">
          <span class="soc-cab"><span class="soc-punto" aria-hidden="true"></span><strong>${esc(S.nombreSociedad(s))}</strong><span class="soc-n">${t.n} ${t.n === 1 ? 'factura' : 'facturas'}</span></span>
          <span class="soc-fila"><span class="soc-et">Colones</span><span class="soc-monto">${money(t.CRC, 'CRC')}</span></span>${barra(t.CRC, maxCRC, 'b-soc')}
          <span class="soc-fila"><span class="soc-et">Dólares</span><span class="soc-monto">${money(t.USD, 'USD')}</span></span>${barra(t.USD, maxUSD, 'b-soc b-claro')}
        </a>`).join('')}</div>
    </section>
    <section class="bloque"><h2>Antigüedad de saldos</h2>
      <div class="tramos">${porTramo.map(({ n, c, t }) => `
        <div class="tramo ${c}">
          <span class="soc-cab"><span class="soc-punto" aria-hidden="true"></span><strong>${n}</strong><span class="soc-n">${t.n} ${t.n === 1 ? 'factura' : 'facturas'}</span></span>
          <span class="soc-fila"><span class="soc-et">Colones</span><span class="soc-monto">${money(t.CRC, 'CRC')}</span></span>${barra(t.CRC, maxTC, 'b-tramo')}
          <span class="soc-fila"><span class="soc-et">Dólares</span><span class="soc-monto">${money(t.USD, 'USD')}</span></span>${barra(t.USD, maxTU, 'b-tramo b-claro')}
        </div>`).join('')}</div>
      <p class="sub">EST: ${pend.filter((f) => f.venceEstimado).length} facturas con saldo usan vencimiento estimado (factura + 30 días).</p>
    </section>
  </div>
  ${seccionAutorizados()}`;
}

// ---------------------------------------------------------------- Reportes

/** "SOC, AGRICOLA COMERCIAL LA HILDA S,A" → "SOC. AGRICOLA COMERCIAL LA HILDA S.A." */
function nombreLegalSociedad(clave) {
  let n = String(clave).replace(/,/g, '.').replace(/\s+/g, ' ').trim();
  if (/S\.A$/.test(n)) n += '.';
  return n;
}

/** Facturas con saldo cuyo vencimiento está en el rango, agrupadas por sociedad y proveedor. */
function datosReporte() {
  const R = ui.rep;
  const socs = R.soc ? [R.soc] : Object.keys(S.sociedades());
  const lista = S.conSaldo().filter((f) => f.moneda === R.mon && socs.includes(f.sociedad) && (!R.prov || f.codigo === R.prov)
    && (!R.desde || (f.fechaVence || '') >= R.desde) && (!R.hasta || (f.fechaVence || '') <= R.hasta));
  return socs.map((soc) => {
    const facts = lista.filter((f) => f.sociedad === soc);
    const porProv = new Map();
    for (const f of facts) {
      if (!porProv.has(f.codigo)) porProv.set(f.codigo, { prov: S.proveedor(f.codigo), facturas: [] });
      porProv.get(f.codigo).facturas.push(f);
    }
    const proveedores = [...porProv.values()].sort((a, b) => a.prov.nombre.localeCompare(b.prov.nombre));
    for (const g of proveedores) {
      g.facturas.sort((a, b) => (a.fechaVence || '').localeCompare(b.fechaVence || '') || String(a.documento).localeCompare(String(b.documento)));
      g.monto = round2(g.facturas.reduce((t, f) => t + f.monto, 0));
      g.saldo = round2(g.facturas.reduce((t, f) => t + S.saldo(f), 0));
    }
    return { soc, proveedores, cantidad: facts.length, monto: round2(proveedores.reduce((t, g) => t + g.monto, 0)), saldo: round2(proveedores.reduce((t, g) => t + g.saldo, 0)) };
  }).filter((x) => R.soc || x.cantidad);
}

/** Días respecto a hoy: negativo si está vencida, positivo si falta para vencer. */
function diasReporte(f) { const d = S.diasVencida(f); return d == null ? null : -d; }
function celdaDias(f) {
  const d = diasReporte(f);
  if (d == null) return '';
  if (d === 0) return '<span class="d-hoy">Hoy</span>';
  return d < 0 ? `<span class="d-venc">−${fmtNum(-d).replace(/,00$/, '')}</span>` : `<span class="d-porv">+${fmtNum(d).replace(/,00$/, '')}</span>`;
}

/** Lista de proveedores tipo filtro de Excel: botón + casilla para escribir + lista. */
function dropProveedor(k, actual, lista, textoTodos) {
  const abierto = ui.provDrop === k;
  return `<div class="rep-prov-drop"><span class="lbl">Proveedor</span>
    <button type="button" class="rep-prov-boton ${actual ? 'elegido' : ''}" data-act="r-prov-abrir" data-v="${k}" aria-expanded="${abierto}">
      <span>${esc(actual ? S.proveedor(actual).nombre : textoTodos)}</span><span aria-hidden="true">▾</span></button>
    ${abierto ? `<div class="rep-prov-panel">
      <input type="search" data-in="r-provfiltro" placeholder="Escriba para buscar…" autocomplete="off" autocorrect="off" autocapitalize="characters" spellcheck="false">
      <ul class="rep-prov-lista">
        ${textoTodos === 'Todos los proveedores' ? `<li><button type="button" data-act="r-prov-set" data-k="${k}" data-v="" class="${actual ? '' : 'on'}">(Todos los proveedores)</button></li>` : ''}
        ${lista.map((p) => `<li data-nombre="${esc(p.nombre.toLowerCase())}"><button type="button" data-act="r-prov-set" data-k="${k}" data-v="${esc(p.codigo)}" class="${actual === p.codigo ? 'on' : ''}">${esc(p.nombre)}</button></li>`).join('')}
      </ul>
      <p class="sub rep-prov-nada" ${lista.length ? 'hidden' : ''}>Ningún proveedor coincide.</p>
    </div>` : ''}
  </div>`;
}

/** Proveedores con saldo en la moneda (y sociedad) elegidas, para el filtro. */
function proveedoresReporte() {
  const R = ui.rep;
  const cods = new Set(S.conSaldo().filter((f) => f.moneda === R.mon && (!R.soc || f.sociedad === R.soc)).map((f) => f.codigo));
  if (R.prov) cods.add(R.prov);
  return [...cods].map((c) => S.proveedor(c)).sort((a, b) => a.nombre.localeCompare(b.nombre));
}

function textoRango() {
  const R = ui.rep;
  if (R.desde && R.hasta) return `Facturas con vencimiento del ${fmtFecha(R.desde)} al ${fmtFecha(R.hasta)}`;
  if (R.hasta) return `Facturas con vencimiento hasta el ${fmtFecha(R.hasta)}`;
  if (R.desde) return `Facturas con vencimiento desde el ${fmtFecha(R.desde)}`;
  return 'Todas las facturas con saldo';
}

function pestanasReportes() {
  const t = ui.rep.tipo;
  return `<div class="rep-pestanas no-imprimir" role="tablist">
    <button class="rep-pestana ${t === 'cxp' ? 'on' : ''}" data-act="r-tipo" data-v="cxp" role="tab" aria-selected="${t === 'cxp'}">Reportes</button>
    <button class="rep-pestana ${t === 'ec' ? 'on' : ''}" data-act="r-tipo" data-v="ec" role="tab" aria-selected="${t === 'ec'}">Estado de cuenta</button>
  </div>`;
}

/** Número para las tablas de reporte: bloque de ancho fijo, centrado bajo el título y alineado por los decimales. */
const cifra = (n) => `<span class="cifra">${fmtNum(n)}</span>`;
/** Total con el signo de la moneda (₡ o $). */
const cifraM = (n, mon) => `<span class="cifra cifra-total">${money(n, mon)}</span>`;

function vistaReportes() { return pestanasReportes() + (ui.rep.tipo === 'ec' ? vistaEstadoCuenta() : vistaCxp()); }

function botonesSalida(prefijo, hay, extrasExcel) {
  const compartir = puedeCompartirArchivos();
  return `<div class="acciones rep-acciones">
      ${compartir ? `<button class="btn primario" data-act="${prefijo}-compartir" ${hay ? '' : 'disabled'}>Compartir PDF</button>` : ''}
      <button class="btn ${compartir ? '' : 'primario'}" data-act="${prefijo}-pdf" ${hay ? '' : 'disabled'}>Descargar PDF</button>
      ${extrasExcel}
      <button class="btn" data-act="r-imprimir" data-v="${prefijo}" ${hay ? '' : 'disabled'}>Imprimir</button></div>`;
}

function vistaCxp() {
  const R = ui.rep;
  const socs = Object.keys(S.sociedades());
  const grupos = datosReporte();
  const emitido = new Date();
  const mon = MONEDAS[R.mon];
  const hoy = hoyISO();
  const preset = (n, d, h) => `<button class="chip ${R.desde === d && R.hasta === h ? 'on' : ''}" data-act="r-preset" data-v="${d}|${h}">${n}</button>`;
  const total = { n: grupos.reduce((t, g) => t + g.cantidad, 0), monto: round2(grupos.reduce((t, g) => t + g.monto, 0)), saldo: round2(grupos.reduce((t, g) => t + g.saldo, 0)) };
  return `
  <div class="titulo-vista no-imprimir"><div><h1>Cuentas por pagar</h1><p class="sub">Facturas con saldo, por fecha de vencimiento. Un reporte por moneda.</p></div>
    ${botonesSalida('r', total.n, `<button class="btn" data-act="r-excel" ${total.n ? '' : 'disabled'} title="Agrupado por proveedor, con totales">Excel reporte</button>
      <button class="btn" data-act="r-excel-lista" ${total.n ? '' : 'disabled'} title="Una fila por factura, con filtros">Excel lista</button>`)}</div>
  <section class="filtros no-imprimir">
    <div class="form-grid rep-filtros">
      <label>Sociedad <select data-in="r-soc"><option value="">Todas las sociedades</option>${socs.map((x) => `<option value="${esc(x)}" ${R.soc === x ? 'selected' : ''}>${esc(S.nombreSociedad(x))}</option>`).join('')}</select></label>
      ${dropProveedor('rep', R.prov, proveedoresReporte(), 'Todos los proveedores')}
      <div><span class="lbl">Moneda</span><div class="chips">
        <button class="chip ${R.mon === 'CRC' ? 'on' : ''}" data-act="r-mon" data-v="CRC">Colones</button>
        <button class="chip ${R.mon === 'USD' ? 'on' : ''}" data-act="r-mon" data-v="USD">Dólares</button></div></div>
      <label>Vence desde <input type="date" data-in="r-desde" value="${R.desde}"></label>
      <label>Vence hasta <input type="date" data-in="r-hasta" value="${R.hasta}"></label>
    </div>
    <div class="chips rep-presets"><span class="lbl">Rápido:</span>
      ${preset('Vencidas a hoy', '', hoy)}${preset('Próximos 7 días', hoy, sumarDias(hoy, 7))}${preset('Próximos 30 días', hoy, sumarDias(hoy, 30))}${preset('Todo', '', '')}
      <button type="button" class="btn chico rep-limpiar" data-act="r-limpiar">Limpiar filtros</button></div>
  </section>
  ${total.n ? '' : '<p class="vacio grande no-imprimir">No hay facturas con saldo en ese rango.</p>'}
  <div class="reporte-hojas ${R.prov ? 'seguido' : ''}">${grupos.filter((g) => g.cantidad).map((g, i) => `
    <article class="reporte">
      ${R.prov && i > 0 ? `<header class="rep-cab rep-cab-corta"><h2 class="rep-soc">${esc(nombreLegalSociedad(g.soc))}</h2></header>` : `<header class="rep-cab">
        <div><h2 class="rep-soc">${esc(nombreLegalSociedad(g.soc))}</h2>
          <p class="rep-titulo">Reporte de Cuentas por Pagar</p>
          <p class="rep-rango">${textoRango()} · ${mon.nombre}</p>
          <p class="rep-leyenda">Días: <span class="d-venc">−</span> vencida hace esos días · <span class="d-porv">+</span> faltan esos días para vencer</p></div>
        <div class="rep-emitido">Emitido<br><strong>${fmtFecha(hoy)}</strong><br>${String(emitido.getHours()).padStart(2, '0')}:${String(emitido.getMinutes()).padStart(2, '0')}</div>
      </header>`}
      ${g.proveedores.map((x) => `
      <section class="rep-prov">
        <h3 class="rep-prov-nombre">${esc(x.prov.nombre)}</h3>
        <table class="rep-tabla">
          <thead><tr><th class="w-doc">N°<span class="larga"> Factura</span></th><th class="c-ff w-ff">Fecha de<br>factura</th><th class="w-fv"><span class="larga">Fecha de<br>vencimiento</span><span class="corta">Vence</span></th><th class="w-dias">Días</th><th class="num w-saldo">Saldo</th><th class="c-mon w-mon">Moneda</th></tr></thead>
          <tbody>${x.facturas.map((f) => `<tr><td>${esc(f.documento)}</td><td class="c-ff">${fmtFecha(f.fechaFactura)}</td><td><span class="larga">${fmtFecha(f.fechaVence)}</span><span class="corta">${fmtFecha(f.fechaVence).replace(/\/20(\d\d)$/, '/$1')}</span></td><td class="c-dias">${celdaDias(f)}</td>
            <td class="num">${cifra(S.saldo(f))}</td><td class="c-mon">${mon.nombre}</td></tr>`).join('')}</tbody>
          ${R.prov ? '' : `<tfoot><tr class="rep-tot-num"><td></td><td class="c-ff"></td><td></td><td class="rep-tot-junto"><span>Total Cliente · ${x.facturas.length} ${x.facturas.length === 1 ? 'factura' : 'facturas'}</span></td><td class="num">${cifraM(x.saldo, R.mon)}</td><td class="c-mon"></td></tr></tfoot>`}
        </table>
      </section>`).join('')}
      <table class="rep-tabla rep-tabla-total rep-total-soc"><tbody><tr><td class="w-doc"></td><td class="c-ff w-ff"></td><td class="w-fv"></td><td class="w-dias rep-tot-junto"><span>Total ${esc(S.nombreSociedad(g.soc))} · ${g.cantidad} ${g.cantidad === 1 ? 'factura' : 'facturas'}</span></td><td class="num w-saldo">${cifraM(g.saldo, R.mon)}</td><td class="c-mon w-mon"></td></tr></tbody></table>
    </article>`).join('')}</div>
  ${!R.soc && grupos.filter((g) => g.cantidad).length > 1 ? `<div class="reporte rep-general"><table class="rep-tabla rep-tabla-total rep-total-soc"><tbody><tr><td class="w-doc"></td><td class="c-ff w-ff"></td><td class="w-fv"></td><td class="w-dias rep-tot-junto"><span>Total todas las sociedades · ${total.n} facturas</span></td><td class="num w-saldo">${cifraM(total.saldo, R.mon)}</td><td class="c-mon w-mon"></td></tr></tbody></table></div>` : ''}`;
}

const limpioArchivo = (t) => String(t).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Za-z0-9 .-]+/g, ' ').replace(/\s+/g, ' ').trim();

/** Nombre del archivo: "Reporte CxP - DON CORLEONE S.A. - COLONES" (con la sociedad si se eligió una). */
function nombreArchivoReporte() {
  const R = ui.rep;
  const partes = ['Reporte CxP'];
  if (R.prov) partes.push(limpioArchivo(S.proveedor(R.prov).nombre).slice(0, 60));
  else if (!R.soc) partes.push('Todas');
  if (R.soc) partes.push(limpioArchivo(S.nombreSociedad(R.soc)));
  partes.push(R.mon === 'USD' ? 'DOLARES' : 'COLONES');
  return partes.join(' - ');
}

async function compartirArchivo(crear) {
  let archivo;
  try { archivo = crear(); } catch (err) { avisar('No se pudo crear el PDF: ' + err.message, 'error'); return; }
  try { await navigator.share({ files: [archivo], title: archivo.name.replace(/\.pdf$/, '') }); }
  catch (err) { if (err.name !== 'AbortError') { descargar(archivo.name, archivo); avisar('No se pudo compartir; se descargó el PDF'); } }
}

function puedeCompartirArchivos() {
  try { return !!(navigator.canShare && navigator.canShare({ files: [new File(['x'], 'x.pdf', { type: 'application/pdf' })] })); } catch { return false; }
}

function pdfReporte() {
  const R = ui.rep;
  const grupos = datosReporte().filter((g) => g.cantidad);
  const ahora = new Date();
  const doc = crearPdfReporte({
    titulo: nombreArchivoReporte(), simbolo: R.mon === 'USD' ? '$' : '¢', grupos, porProveedor: !!R.prov, rango: textoRango(), moneda: MONEDAS[R.mon].nombre,
    emitido: { fecha: fmtFecha(hoyISO()), hora: `${String(ahora.getHours()).padStart(2, '0')}:${String(ahora.getMinutes()).padStart(2, '0')}` },
    nombreSociedad: nombreLegalSociedad, nombreCorto: (c) => S.nombreSociedad(c), fmtNum, fmtFecha, dias: diasReporte, saldo: (f) => S.saldo(f),
    total: !R.soc && grupos.length > 1 ? { n: grupos.reduce((t, g) => t + g.cantidad, 0), saldo: round2(grupos.reduce((t, g) => t + g.saldo, 0)) } : null,
  });
  return new File([doc.output('blob')], nombreArchivoReporte() + '.pdf', { type: 'application/pdf' });
}

const fechaExcel = (iso) => { if (!iso) return ''; const [y, m, d] = iso.split('-').map(Number); return new Date(y, m - 1, d); };

function guardarExcel(filas, anchos, nombre, formatos = {}, filtro = null, moneda = null) {
  const XLSX = window.XLSX;
  const ws = XLSX.utils.aoa_to_sheet(filas, { cellDates: true });
  ws['!cols'] = anchos.map((wch) => ({ wch }));
  const rango = XLSX.utils.decode_range(ws['!ref']);
  for (let r = rango.s.r; r <= rango.e.r; r++) {
    for (let c = rango.s.c; c <= rango.e.c; c++) {
      const celda = ws[XLSX.utils.encode_cell({ r, c })];
      if (!celda) continue;
      if (celda.t === 'd' || celda.v instanceof Date) celda.z = 'dd/mm/yyyy';
      else if (celda.t === 'n' && moneda && /^Total/.test(String(filas[r]?.find((x) => typeof x === 'string' && /^Total/.test(x)) || ''))) celda.z = `"${moneda}"#,##0.00`;
      else if (celda.t === 'n' && formatos[c]) celda.z = formatos[c];
    }
  }
  if (filtro) ws['!autofilter'] = { ref: filtro };
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Reporte');
  const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx', cellDates: true });
  descargar(nombre + '.xlsx', new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
}

/** Excel "lista": una fila por factura con todas sus columnas, para filtrar u ordenar en Excel. */
function excelListaReporte() {
  const R = ui.rep;
  const filas = [['Proveedor', 'N° Factura', 'Fecha de vencimiento', 'Días', 'Saldo', 'Moneda', 'Sociedad (origen de la factura)']];
  for (const g of datosReporte().filter((x) => x.cantidad)) {
    for (const x of g.proveedores) {
      for (const f of x.facturas) filas.push([x.prov.nombre, String(f.documento), fechaExcel(f.fechaVence), diasReporte(f) ?? '', round2(S.saldo(f)), MONEDAS[R.mon].nombre, nombreLegalSociedad(f.sociedad)]);
    }
  }
  guardarExcel(filas, [44, 14, 20, 8, 16, 10, 40], nombreArchivoReporte() + ' (lista)', { 3: '+0;-0;"Hoy"', 4: '#,##0.00' }, `A1:G${filas.length}`);
}

function excelReporte() {
  const R = ui.rep;
  const filas = [];
  const hoy = hoyISO();
  for (const g of datosReporte().filter((x) => x.cantidad)) {
    filas.push([nombreLegalSociedad(g.soc)], ['Reporte de Cuentas por Pagar'], [`${textoRango()} · ${MONEDAS[R.mon].nombre}`], [`Emitido: ${fmtFecha(hoy)}`], []);
    for (const x of g.proveedores) {
      filas.push([x.prov.nombre]);
      filas.push(['N° Factura', 'Fecha de factura', 'Fecha de vencimiento', 'Días (− vencida / + por vencer)', 'Saldo', 'Moneda']);
      for (const f of x.facturas) filas.push([String(f.documento), fmtFecha(f.fechaFactura), fmtFecha(f.fechaVence), diasReporte(f) ?? '', round2(S.saldo(f)), MONEDAS[R.mon].nombre]);
      if (!R.prov) filas.push(['', '', '', `Total Cliente · ${x.facturas.length} ${x.facturas.length === 1 ? 'factura' : 'facturas'}`, x.saldo, '']);
      filas.push([]);
    }
    filas.push(['', '', '', `Total ${S.nombreSociedad(g.soc)} · ${g.cantidad} ${g.cantidad === 1 ? 'factura' : 'facturas'}`, g.saldo, ''], [], []);
  }
  const conDatos = datosReporte().filter((x) => x.cantidad);
  if (!R.soc && conDatos.length > 1) {
    filas.push(['', '', '', `Total todas las sociedades · ${conDatos.reduce((t, g) => t + g.cantidad, 0)} facturas`, round2(conDatos.reduce((t, g) => t + g.saldo, 0)), '']);
  }
  const XLSX = window.XLSX;
  const ws = XLSX.utils.aoa_to_sheet(filas);
  ws['!cols'] = [{ wch: 44 }, { wch: 16 }, { wch: 20 }, { wch: 34 }, { wch: 18 }, { wch: 10 }];
  // Formato: montos con miles y céntimos; días con signo (+ por vencer, − vencida)
  const rango = XLSX.utils.decode_range(ws['!ref']);
  for (let r = rango.s.r; r <= rango.e.r; r++) {
    for (const c of [3, 4]) {
      const celda = ws[XLSX.utils.encode_cell({ r, c })];
      const esTotal = /^Total/.test(String(filas[r]?.[3] || ''));
      if (celda && celda.t === 'n') celda.z = c === 3 ? '+0;-0;"Hoy"' : esTotal ? `"${MONEDAS[R.mon].simbolo}"#,##0.00` : '#,##0.00';
    }
  }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Reporte');
  const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  descargar(nombreArchivoReporte() + '.xlsx', new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
}

// ---------------------------------------------------------------- Estado de cuenta

const TIPOS_MOV = { pago: 'Pago', abono: 'Abono', nc: 'Nota de crédito', cargo: 'Intereses / otros cargos' };

/** Proveedores con facturas en la moneda (y sociedad) elegidas. */
function proveedoresEstado() {
  const E = ui.ec;
  const base = E.alcance === 'saldo' ? S.conSaldo() : S.listaFacturas();
  const cods = new Set(base.filter((f) => f.moneda === E.mon && (!E.soc || f.sociedad === E.soc)).map((f) => f.codigo));
  if (E.prov) cods.add(E.prov);
  return [...cods].map((c) => S.proveedor(c)).sort((a, b) => a.nombre.localeCompare(b.nombre));
}

/** Por sociedad: cada factura con su monto y, debajo, sus movimientos con el saldo después de cada uno. */
function datosEstadoCuenta() {
  const E = ui.ec;
  if (!E.prov) return [];
  const socs = E.soc ? [E.soc] : Object.keys(S.sociedades());
  const facts = S.listaFacturas().filter((f) => f.codigo === E.prov && f.moneda === E.mon && socs.includes(f.sociedad)
    && (E.alcance !== 'saldo' || S.saldo(f) > 0.004)
    && (!E.desde || (f.fechaFactura || '') >= E.desde) && (!E.hasta || (f.fechaFactura || '') <= E.hasta));
  return socs.map((soc) => {
    const lista = facts.filter((f) => f.sociedad === soc)
      .sort((a, b) => (a.fechaFactura || '').localeCompare(b.fechaFactura || '') || String(a.documento).localeCompare(String(b.documento), 'es', { numeric: true }));
    const t = { facturado: 0, abonos: 0, nc: 0, cargos: 0, saldo: 0 };
    const facturas = lista.map((f) => {
      let saldo = round2(f.monto);
      const lineas = [{ fecha: f.fechaFactura, tipo: 'factura', detalle: `Factura ${f.documento}` + (f.fechaVence ? ` · vence ${fmtFecha(f.fechaVence)}` : ''), cargo: round2(f.monto), abono: null, saldo }];
      const movs = [...S.movimientosDe(f)].sort((a, b) => (a.fecha || '9999').localeCompare(b.fecha || '9999'));
      for (const m of movs) {
        const monto = round2(m.monto);
        const comp = m.comprobante ? (m.tipo === 'nc' ? ` N° ${m.comprobante}` : ` · Comp. ${m.comprobante}`) : '';
        const ref = m.tipo === 'nc' ? `aplicada a la factura ${f.documento}` : `a la factura ${f.documento}`;
        if (m.tipo === 'cargo') {
          t.cargos = round2(t.cargos + monto);
          lineas.push({ fecha: m.fecha, tipo: 'cargo', detalle: TIPOS_MOV.cargo + ' (pagados)' + comp, ref, cargo: monto, abono: monto, saldo, arch: S.comprobanteDe(m) });
        } else {
          saldo = round2(saldo - monto);
          if (m.tipo === 'nc') t.nc = round2(t.nc + monto); else t.abonos = round2(t.abonos + monto);
          lineas.push({ fecha: m.fecha, tipo: m.tipo, detalle: (TIPOS_MOV[m.tipo] || m.tipo) + comp, ref, cargo: null, abono: monto, saldo, arch: S.comprobanteDe(m) });
        }
      }
      t.facturado = round2(t.facturado + f.monto);
      t.saldo = round2(t.saldo + S.saldo(f));
      return { f, lineas };
    });
    return { soc, facturas, t };
  }).filter((g) => g.facturas.length);
}

function textoRangoEstado() {
  const E = ui.ec;
  const alc = E.alcance === 'saldo' ? 'Facturas con saldo' : 'Todas las facturas';
  if (E.desde && E.hasta) return `${alc} del ${fmtFecha(E.desde)} al ${fmtFecha(E.hasta)} (fecha de factura)`;
  if (E.desde) return `${alc} desde el ${fmtFecha(E.desde)}`;
  if (E.hasta) return `${alc} hasta el ${fmtFecha(E.hasta)}`;
  return alc;
}

function totalesEstado(grupos) {
  const z = { facturado: 0, abonos: 0, nc: 0, cargos: 0, saldo: 0 };
  for (const g of grupos) for (const k in z) z[k] = round2(z[k] + g.t[k]);
  return z;
}

function nombreArchivoEstado() {
  const E = ui.ec;
  const partes = ['Estado de cuenta', limpioArchivo(S.proveedor(E.prov).nombre).slice(0, 60)];
  if (E.soc) partes.push(limpioArchivo(S.nombreSociedad(E.soc)));
  partes.push(E.mon === 'USD' ? 'DOLARES' : 'COLONES');
  return partes.join(' - ');
}

function vistaEstadoCuenta() {
  const E = ui.ec;
  const socs = Object.keys(S.sociedades());
  const grupos = datosEstadoCuenta();
  const z = totalesEstado(grupos);
  const hay = grupos.length > 0;
  const ahora = new Date();
  const fechaCorta = (iso) => (iso ? `<span class="larga">${fmtFecha(iso)}</span><span class="corta">${fmtFecha(iso).replace(/\/20(\d\d)$/, '/$1')}</span>` : '<span class="sub">sin fecha</span>');
  const filaTotales = (etq, t, cls = '') => `<tr class="ec-tot ${cls}"><td colspan="2" class="ec-tot-lbl">${etq}</td><td class="num">${cifraM(t.facturado + t.cargos, E.mon)}</td><td class="num">${cifraM(t.abonos + t.nc + t.cargos, E.mon)}</td><td class="num">${cifraM(t.saldo, E.mon)}</td></tr>`;
  return `
  <div class="titulo-vista no-imprimir"><div><h1>Estado de cuenta</h1><p class="sub">Todos los movimientos de un proveedor: cada factura con su monto, abonos, pagos, notas de crédito y el saldo.</p></div>
    ${botonesSalida('ec', hay, `<button class="btn" data-act="ec-excel" ${hay ? '' : 'disabled'}>Descargar Excel</button>`)}</div>
  <section class="filtros no-imprimir">
    <div class="form-grid rep-filtros">
      ${dropProveedor('ec', E.prov, proveedoresEstado(), 'Elija un proveedor')}
      <label>Sociedad <select data-in="ec-soc"><option value="">Todas las sociedades</option>${socs.map((x) => `<option value="${esc(x)}" ${E.soc === x ? 'selected' : ''}>${esc(S.nombreSociedad(x))}</option>`).join('')}</select></label>
      <div><span class="lbl">Moneda</span><div class="chips">
        <button class="chip ${E.mon === 'CRC' ? 'on' : ''}" data-act="ec-mon" data-v="CRC">Colones</button>
        <button class="chip ${E.mon === 'USD' ? 'on' : ''}" data-act="ec-mon" data-v="USD">Dólares</button></div></div>
      <div><span class="lbl">Facturas</span><div class="chips">
        <button class="chip ${E.alcance === 'saldo' ? 'on' : ''}" data-act="ec-alcance" data-v="saldo">Con saldo</button>
        <button class="chip ${E.alcance === 'todas' ? 'on' : ''}" data-act="ec-alcance" data-v="todas">Todas</button></div></div>
      <label>Fecha de factura desde <input type="date" data-in="ec-desde" value="${E.desde}"></label>
      <label>Hasta <input type="date" data-in="ec-hasta" value="${E.hasta}"></label>
    </div>
    <div class="chips rep-presets"><button type="button" class="btn chico rep-limpiar" data-act="ec-limpiar">Limpiar filtros</button></div>
  </section>
  ${!E.prov ? '<p class="vacio grande no-imprimir">Elija un proveedor para ver su estado de cuenta.</p>' : !hay ? '<p class="vacio grande no-imprimir">Ese proveedor no tiene facturas con esos filtros.</p>' : `
  <div class="reporte-hojas seguido">
    <article class="reporte estado-cuenta">
      <header class="rep-cab">
        <div><h2 class="rep-soc">${esc(S.proveedor(E.prov).nombre)}</h2>
          <p class="rep-titulo">Estado de Cuenta</p>
          <p class="rep-rango">${E.soc ? esc(nombreLegalSociedad(E.soc)) : 'Todas las sociedades'} · ${MONEDAS[E.mon].nombre}</p>
          <p class="rep-rango">${textoRangoEstado()}</p></div>
        <div class="rep-emitido">Emitido<br><strong>${fmtFecha(hoyISO())}</strong><br>${String(ahora.getHours()).padStart(2, '0')}:${String(ahora.getMinutes()).padStart(2, '0')}</div>
      </header>
      ${grupos.map((g) => `
      <section class="ec-soc">
        <h3 class="ec-soc-nombre">${esc(nombreLegalSociedad(g.soc))}</h3>
        <table class="rep-tabla ec-tabla">
          <thead><tr><th class="ec-w-fecha">Fecha</th><th class="ec-w-det">Detalle</th><th class="ec-w-num">Monto</th><th class="ec-w-num">Abonos</th><th class="ec-w-num">Saldo</th></tr></thead>
          ${g.facturas.map(({ lineas }) => `<tbody class="ec-factura">${lineas.map((l) => `<tr class="ec-${l.tipo === 'factura' ? 'fact' : 'mov'} ${l.tipo === 'nc' ? 'ec-nc' : ''}">
            <td>${fechaCorta(l.fecha)}</td><td class="ec-det">${l.tipo === 'factura' ? esc(l.detalle).replace(' · vence ', '<span class="ec-vence"> · vence ').replace(/(vence [0-9/]+)$/, '$1</span>') : `<span class="ec-flecha">↳</span> ${esc(l.detalle)} <span class="ec-ref">· ${esc(l.ref)}</span>${linkComprobante(l.arch)}`}</td>
            <td class="num">${l.cargo == null ? '' : cifra(l.cargo)}</td><td class="num">${l.abono == null ? '' : cifra(l.abono)}</td><td class="num">${cifra(l.saldo)}</td></tr>`).join('')}</tbody>`).join('')}
          <tfoot>${filaTotales(`Total ${esc(S.nombreSociedad(g.soc))} · ${g.facturas.length} ${g.facturas.length === 1 ? 'factura' : 'facturas'}`, g.t)}</tfoot>
        </table>
        <p class="ec-resumen">Facturado ${fmtNum(g.t.facturado)} · Abonos y pagos ${fmtNum(g.t.abonos)} · Notas de crédito ${fmtNum(g.t.nc)}${g.t.cargos ? ` · Intereses pagados ${fmtNum(g.t.cargos)}` : ''} · <strong>Saldo ${fmtNum(g.t.saldo)}</strong></p>
      </section>`).join('')}
      ${grupos.length > 1 ? `<table class="rep-tabla ec-tabla ec-general"><colgroup><col class="ec-w-fecha"><col class="ec-w-det"><col class="ec-w-num"><col class="ec-w-num"><col class="ec-w-num"></colgroup><tbody>${filaTotales(`Total todas las sociedades · ${grupos.reduce((n, g) => n + g.facturas.length, 0)} facturas`, z, 'ec-tot-general')}</tbody></table>` : ''}
    </article>
  </div>`}`;
}

function excelEstado() {
  const E = ui.ec;
  const grupos = datosEstadoCuenta();
  const filas = [[S.proveedor(E.prov).nombre], ['Estado de Cuenta'], [`${E.soc ? nombreLegalSociedad(E.soc) : 'Todas las sociedades'} · ${MONEDAS[E.mon].nombre}`], [textoRangoEstado()], [`Emitido: ${fmtFecha(hoyISO())}`], [],
    ['Sociedad', 'Factura', 'Fecha', 'Detalle', 'Monto', 'Abonos', 'Saldo']];
  for (const g of grupos) {
    for (const { f, lineas } of g.facturas) {
      for (const l of lineas) filas.push([S.nombreSociedad(g.soc), String(f.documento), fechaExcel(l.fecha), l.tipo === 'factura' ? l.detalle : `   ↳ ${l.detalle} · ${l.ref}`, l.cargo ?? '', l.abono ?? '', l.saldo]);
    }
    filas.push(['', '', '', `Total ${S.nombreSociedad(g.soc)} · ${g.facturas.length} facturas`, round2(g.t.facturado + g.t.cargos), round2(g.t.abonos + g.t.nc + g.t.cargos), g.t.saldo], []);
  }
  if (grupos.length > 1) { const z = totalesEstado(grupos); filas.push(['', '', '', 'Total todas las sociedades', round2(z.facturado + z.cargos), round2(z.abonos + z.nc + z.cargos), z.saldo]); }
  guardarExcel(filas, [18, 12, 12, 60, 16, 16, 16], nombreArchivoEstado(), { 4: '#,##0.00', 5: '#,##0.00', 6: '#,##0.00' }, null, MONEDAS[E.mon].simbolo);
}

function pdfEstado() {
  const E = ui.ec;
  const grupos = datosEstadoCuenta();
  const ahora = new Date();
  const doc = crearPdfEstadoCuenta({
    titulo: nombreArchivoEstado(), simbolo: E.mon === 'USD' ? '$' : '¢', proveedor: S.proveedor(E.prov).nombre, grupos,
    subtitulo: `${E.soc ? nombreLegalSociedad(E.soc) : 'Todas las sociedades'} · ${MONEDAS[E.mon].nombre}`, rango: textoRangoEstado(),
    emitido: { fecha: fmtFecha(hoyISO()), hora: `${String(ahora.getHours()).padStart(2, '0')}:${String(ahora.getMinutes()).padStart(2, '0')}` },
    nombreSociedad: nombreLegalSociedad, nombreCorto: (c) => S.nombreSociedad(c), fmtNum, fmtFecha,
    total: grupos.length > 1 ? { n: grupos.reduce((n, g) => n + g.facturas.length, 0), ...totalesEstado(grupos) } : null,
  });
  return new File([doc.output('blob')], nombreArchivoEstado() + '.pdf', { type: 'application/pdf' });
}

// ---------------------------------------------------------------- Configuración

function seccionUsuarios() {
  if (!Sync.gestionUsuarios()) {
    return `<section class="bloque">
      <h2>Usuarios</h2>
      <p class="sub">Los usuarios se administran en la pestaña <strong>Usuarios</strong> de la hoja de Google:</p>
      <ul class="lista-pasos">
        <li><strong>Agregar a alguien:</strong> fila nueva con usuario (en minúsculas), nombre, rol (<span class="mono">admin</span>, <span class="mono">aplica</span> o <span class="mono">aprueba</span>), <span class="mono">TRUE</span> y la palabra <span class="mono">${esc(Sync.CLAVE_INICIAL)}</span> en <em>contrasenaTemporal</em>. Para que aparezca en la pantalla de inicio, avíseme y lo agrego a la lista de la aplicación.</li>
        <li><strong>Si alguien no puede entrar:</strong> escriba <span class="mono">${esc(Sync.CLAVE_INICIAL)}</span> en su columna <em>contrasenaTemporal</em>.</li>
        <li><strong>Quitar el acceso:</strong> ponga <span class="mono">FALSE</span> en <em>activo</em>.</li>
      </ul>
      <p class="sub">Roles: <strong>admin</strong> ve todo, incluida Configuración completa; <strong>aplica</strong> registra facturas y aplica pagos; <strong>aprueba</strong> programa pagos.</p>
    </section>`;
  }
  const lista = Sync.usuariosConocidos() || [];
  return `<section class="bloque">
    <h2>Usuarios</h2>
    <p class="sub">Cada persona entra tocando su nombre en la pantalla de inicio.</p>
    <div class="tabla-scroll"><table class="tabla">
      <thead><tr><th>Usuario</th><th>Nombre</th><th>Rol</th><th>Estado</th><th></th></tr></thead>
      <tbody>${lista.map((u) => `<tr>
        <td class="mono">${esc(u.usuario)}</td><td>${esc(u.nombre)}</td><td>${esc(ROLES[u.rol] || u.rol)}</td>
        <td>${u.activo === false ? '<span class="estado rev">Desactivado</span>' : u.pendiente ? '<span class="estado abon">Con PIN temporal</span>' : '<span class="estado canc">Activo</span>'}</td>
        <td class="nowrap"><button class="btn chico" data-act="u-pin" data-v="${esc(u.usuario)}">PIN temporal nuevo</button>
          ${u.usuario === S.usuarioActual()?.usuario ? '' : `<button class="btn chico" data-act="u-activo" data-v="${esc(u.usuario)}" data-activo="${u.activo === false ? '1' : '0'}">${u.activo === false ? 'Activar' : 'Desactivar'}</button>`}</td>
      </tr>`).join('') || '<tr><td colspan="5" class="sub">Sincronice para ver la lista.</td></tr>'}</tbody>
    </table></div>
    <h3>Agregar usuario</h3>
    <form class="form-grid" data-form="u-crear">
      <label>Usuario <input name="usuario" required placeholder="Ej. fernando" autocapitalize="none" spellcheck="false" pattern="[a-z][a-z0-9._-]{1,29}"></label>
      <label>Nombre completo <input name="nombre" required placeholder="Ej. Fernando Pérez"></label>
      <label>Rol <select name="rol"><option value="admin">Administra y aplica</option><option value="aplica">Aplica pagos</option><option value="aprueba">Aprueba pagos</option></select></label>
      <div class="ancho"><button class="btn primario" type="submit">Crear usuario</button></div>
    </form>
  </section>`;
}

function mostrarPin(nombre, pin) {
  abrirDialogo(`
    <header class="dlg-cab"><div><h2>PIN temporal de ${esc(nombre)}</h2></div><button class="btn-cerrar" data-act="cerrar" aria-label="Cerrar">×</button></header>
    <div class="dlg-cuerpo centro-txt">
      <p class="pin-grande">${esc(pin)}</p>
      <p class="sub">Entrégueselo en privado. Al entrar con este PIN, la aplicación le pedirá crear el suyo. También queda anotado en la pestaña Usuarios de la hoja hasta que lo cambie.</p>
      <button class="btn primario" data-act="cerrar">Listo</button>
    </div>`);
}

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
    </div>
    ${S.esAdmin() ? `<h3>Conexión de este dispositivo</h3>
    <p class="sub">${Sync.conexion() ? 'Conectado. La dirección y la contraseña del servicio se ven con la contraseña del servicio.' : 'Sin conexión configurada.'}</p>
    <button class="btn" data-act="conexion">Ver o cambiar la conexión</button>` : ''}
    <p class="sub">${e.pendientes.length ? `${e.pendientes.length} cambios esperando conexión para enviarse.` : 'No hay cambios pendientes de enviar.'}${e.error ? ' Último error: ' + esc(e.error) : ''}</p>
  </section>

  <section class="bloque">
    <h2>Banco</h2>
    <form class="form-grid" data-form="cfg-banco">
      <label>Cédula del empleado que sube el archivo <input name="controlEncabezado" inputmode="numeric" value="${esc(String(S.cfg('controlEncabezado')).replace(/^0+/, ''))}" placeholder="Ej. 205470116"></label>
      <p class="ancho sub">Es la «Cédula Empleado» de la plantilla del BCR (hoy la de Daniel). Va en todos los archivos, de todas las sociedades. La «Cédula jurídica», la cuenta y el consecutivo de cada sociedad están más abajo.</p>
      <div class="ancho"><button class="btn primario" type="submit">Guardar</button></div>
    </form>
  </section>

  <section class="bloque">
    <h2>Sociedades: datos para el archivo del BCR</h2>
    ${socs.length ? socs.map(([k, s]) => `
    <form class="form-grid sociedad-form" data-form="cfg-soc" data-v="${esc(k)}">
      <h3 class="ancho">${esc(k)}</h3>
      <label>Nombre corto <input name="corto" value="${esc(s.corto)}"></label>
      <label>Descripción en la línea de débito <input name="descripcionBanco" value="${esc(s.descripcionBanco)}" placeholder="Ej. PROVEE ${esc(String(s.corto).toUpperCase())}"></label>
      <label>IBAN colones <input name="cuentaCRC" class="mono" value="${esc(fmtIBAN(s.cuentaCRC))}" data-in="iban"><span class="ayuda" data-ayuda></span></label>
      <label>IBAN dólares <input name="cuentaUSD" class="mono" value="${esc(fmtIBAN(s.cuentaUSD))}" data-in="iban"><span class="ayuda" data-ayuda></span></label>
      <label>Cédula jurídica (como en la plantilla del BCR) <input name="cedulaBCR" inputmode="numeric" value="${esc(S.bancoSoc(k).cedula || '')}" placeholder="Ej. 6133"></label>
      <label>Último consecutivo usado <input name="ultimoConsec" inputmode="numeric" maxlength="3" value="${esc(S.bancoSoc(k).ultimo ?? '')}" placeholder="Ej. 5">
        <span class="sub">El próximo archivo usará el ${S.siguienteConsecutivo(k)}.</span></label>
      <div class="ancho"><button class="btn" type="submit">Guardar ${esc(s.corto)}</button></div>
    </form>`).join('') : '<p class="sub">Aparecerán al cargar el Excel.</p>'}
  </section>

  ${S.esAdmin() ? seccionUsuarios() : ''}

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
  etiquetarTablas($dialog);
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

/** Copia el título de cada columna a sus celdas: en el celular cada fila se ve como tarjeta. */
function etiquetarTablas(raiz = document) {
  raiz.querySelectorAll('table.tabla').forEach((t) => {
    const titulos = [...t.querySelectorAll('thead th')].map((th) => th.textContent.trim());
    t.querySelectorAll('tbody tr').forEach((tr) => {
      let col = 0;
      [...tr.children].forEach((td) => {
        if (!td.hasAttribute('data-label')) td.setAttribute('data-label', titulos[col] || '');
        col += Number(td.getAttribute('colspan') || 1);
      });
    });
  });
}

function despuesDeRender(ruta) {
  etiquetarTablas($app);
  if (ruta === 'programar') actualizarBarra();
  if (ruta === 'reportes') { const f = document.querySelector('[data-in="r-provfiltro"]'); if (f?.value) f.dispatchEvent(new Event('input', { bubbles: true })); }
  if (ruta === 'config') document.querySelectorAll('[data-in="iban"]').forEach(validarCampoIBAN);
}

document.addEventListener('click', (e) => {
  if (ui.provDrop && !e.target.closest('.rep-prov-drop')) { ui.provDrop = null; render(); }
  const el = e.target.closest('[data-act]');
  if (!el) return;
  const act = el.dataset.act, v = el.dataset.v;
  if (el.tagName === 'A' || el.tagName === 'TR') e.preventDefault();
  acciones[act]?.(v, el, e);
});

document.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && e.target.matches('tr[data-act]')) acciones[e.target.dataset.act]?.(e.target.dataset.v);
  // «Buscar» en el teclado del celular: muestra los resultados ya y cierra el teclado
  if (e.key === 'Enter' && e.target.matches('[data-in="r-provfiltro"]')) {
    e.preventDefault();
    const vis = [...document.querySelectorAll('.rep-prov-lista li[data-nombre]:not([hidden]) button')];
    if (vis.length === 1) vis[0].click();
    return;
  }
  if (e.key === 'Escape' && ui.provDrop) { ui.provDrop = null; render(); return; }
  if (e.key === 'Enter' && e.target.matches('input[type="search"]')) { e.preventDefault(); clearTimeout(tBusqueda); render(); e.target.blur(); }
});

const acciones = {
  'sync-ahora': () => Sync.sincronizar(),
  'p-revisar': (v) => {
    const f = S.factura(v);
    if (f.revision) resultado(S.actualizarFactura(v, { revision: false }), 'Se quitó la revisión');
    else dialogoPedirRevision(v);
  },
  revisado: (v) => {
    const f = S.factura(v);
    resultado(S.actualizarFactura(v, { revision: false }), S.programacionDe(f) ? 'Revisada: ya está lista para pagar en «Por aplicar»' : 'Revisada');
  },
  'rev-devolver': (v) => dialogoDevolverRevision(v),
  'c-soc': (v) => { ui.canc.soc = v; ui.canc.limite = 100; render(); },
  'c-mas': () => { ui.canc.limite += 300; render(); },
  actualizar: () => actualizarApp(),
  'u-pin': (v) => {
    if (!confirm(`¿Generar un PIN temporal nuevo para ${v}? Su PIN actual dejará de funcionar cuando entre con el nuevo.`)) return;
    Sync.adminUsuarios({ op: 'pin', usuario: v }).then((r) => mostrarPin(v, r.pin)).catch((e) => avisar(e.message, 'error'));
  },
  'u-activo': (v, el) => {
    const activar = el.dataset.activo === '1';
    if (!confirm(activar ? `¿Activar a ${v}?` : `¿Desactivar a ${v}? No podrá entrar hasta que se active de nuevo.`)) return;
    Sync.adminUsuarios({ op: 'activo', usuario: v, activo: activar }).then(() => avisar(activar ? 'Usuario activado' : 'Usuario desactivado')).catch((e) => avisar(e.message, 'error'));
  },
  conexion: () => dialogoConexion(),
  elegir: (v) => {
    if (!v) { ui.login = ''; render(); return; }
    if (!Sync.conexion()) { ui.login = v; render(); return; }
    entrarComo(v);
  },
  'cambiar-usuario': () => {
    try { Sync.cambiarUsuario(); cerrarDialogo(); ui.login = ''; location.hash = ''; render(); } catch (e) { avisar(e.message, 'error'); }
  },
  cuenta: () => acciones['cambiar-usuario'](),
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
    const items = [...ui.prog.sel].map(([facturaId, txt]) => ({ facturaId, monto: parseMonto(txt) }));
    const r = S.programarPagos(items, ui.prog.fecha);
    if (r.error) { avisar(r.error, 'error'); return; }
    ui.prog.sel.clear();
    render();
    avisar(`${items.length} ${items.length === 1 ? 'pago programado' : 'pagos programados'} para el ${fmtFecha(ui.prog.fecha)}` + (r.avisos.length ? ' · ' + r.avisos.join(' · ') : ''));
  },

  'a-devolver': (v) => { if (confirm('¿Devolver esta factura a Norman? Saldrá de este lote.')) resultado(S.quitarProgramacion(v, 'Devuelto'), 'Devuelta a Norman'); },
  'a-forma': (v) => { const i = v.lastIndexOf('|'); ui.forma[decodeURIComponent(v.slice(0, i))] = v.slice(i + 1); render(); },
  'a-txt': (v) => {
    const l = loteDe(v);
    if (!l) return;
    const { banco } = partesLote(l);
    const sub = { ...l, lineas: banco, total: round2(banco.reduce((t, x) => t + x.total, 0)) };
    try {
      const prob = S.problemasLote(sub);
      if (prob.length) { avisar(prob.join(' '), 'error'); return; }
      const n = leerConsecutivo(document.getElementById('consec-' + v));
      if (!n) return;
      descargarTxt(sub, n);
      const r = S.marcarEnviado(banco.flatMap((ln) => ln.items.map(({ p }) => p.id)), n);
      if (resultado(r, `Archivo descargado (consecutivo ${c3(n)}) · quedó «Enviado al banco · esperando comprobante»`)) render();
    } catch (err) { avisar(err.message, 'error'); }
  },
  'a-txt-otra': (v, el) => {
    const l = loteDe(v);
    if (!l) return;
    const { enviadas } = partesLote(l, el.dataset.g), e0 = enviadas[0]?.e;
    if (!enviadas.length) return;
    if (!confirm(`Este archivo ya lo generó ${e0?.por || 'alguien'} (${fechaHora(Date.parse(e0?.fecha))}).\n\nSi ya se subió al banco, subirlo otra vez PAGARÍA DOS VECES.\n\n¿Descargarlo de nuevo de todos modos?`)) return;
    try { descargarTxt({ ...l, lineas: enviadas.map((x) => x.ln) }, e0?.consecutivo || S.siguienteConsecutivo(l.sociedad)); } catch (err) { avisar(err.message, 'error'); }
  },
  'a-txt-rechazo': (v, el) => {
    const l = loteDe(v);
    if (!l) return;
    const { enviadas } = partesLote(l, el.dataset.g), e0 = enviadas[0]?.e;
    if (!enviadas.length) return;
    const n = S.siguienteConsecutivo(l.sociedad);
    if (!confirm(`Se generará el mismo archivo con el consecutivo ${c3(n)}${e0?.consecutivo ? ` (en lugar del ${c3(e0.consecutivo)})` : ''}.\n\nÚselo solo si el banco rechazó el archivo por consecutivo repetido. ¿Continuar?`)) return;
    try {
      descargarTxt({ ...l, lineas: enviadas.map((x) => x.ln) }, n);
      if (resultado(S.marcarEnviado(enviadas.flatMap(({ ln }) => ln.items.map(({ p }) => p.id)), n), `Archivo descargado con el consecutivo ${c3(n)}`)) render();
    } catch (err) { avisar(err.message, 'error'); }
  },
  'a-volver': (v, el) => {
    const l = loteDe(v);
    if (!l) return;
    const { enviadas } = partesLote(l, el.dataset.g);
    if (!confirm('¿Volver a pendiente? Úselo solo si el banco rechazó el archivo o no se llegó a subir. Después podrá generar el archivo otra vez.')) return;
    if (resultado(S.volverPendiente(enviadas.flatMap(({ ln }) => ln.items.map(({ p }) => p.id))), 'Volvió a pendiente')) render();
  },
  'a-ver-txt': (v) => {
    const l = loteDe(v);
    try {
      const { banco } = partesLote(l);
      const txt = textoArchivo({ ...l, lineas: banco }, Number(document.getElementById('consec-' + v)?.value) || S.siguienteConsecutivo(l.sociedad));
      abrirDialogo(`<header class="dlg-cab"><h2>Contenido del archivo</h2><button class="btn-cerrar" data-act="cerrar" aria-label="Cerrar">×</button></header>
        <div class="dlg-cuerpo"><pre class="txt">${esc(txt.replace(/ +\r\n/g, '\r\n'))}</pre><p class="sub">Los espacios al final de cada línea no se muestran; el archivo descargado sí los lleva (214 caracteres por línea).</p></div>`);
    } catch (err) { avisar(err.message, 'error'); }
  },
  'a-aplicar': (v, el) => {
    const l = loteDe(v), g = el.dataset.g;
    if (!l) return;
    const { enviadas } = partesLote(l, g);
    const lineas = enviadas.map((x) => x.ln);
    const sub = { ...l, lineas, total: round2(lineas.reduce((t, x) => t + x.total, 0)) };
    const comp = document.getElementById(`comp-${v}-${g}`).value.trim();
    const fecha = document.getElementById(`fecha-${v}-${g}`).value;
    if (!comp) { avisar('Escriba el comprobante general del débito', 'error'); return; }
    if (!confirm(`¿Aplicar el pago de ${money(sub.total, l.moneda)} (${lineas.length} ${lineas.length === 1 ? 'proveedor' : 'proveedores'}) con el comprobante ${comp}?`)) return;
    const r = S.aplicarLote(sub, { comprobante: comp, fecha });
    if (!resultado(r, 'Pago aplicado')) return;
    ui.apAbierto[r.loteId] = true;
    render();
    adjuntarDespues('gen|' + l.clave + '|' + g, `Debito ${comp}`, (a) => S.guardarAdjLote(r.loteId, a));
  },
  'a-aplicar-manual': (v) => {
    const k = decodeURIComponent(v), i = k.lastIndexOf('|');
    const l = S.lotesPendientes().find((x) => x.clave === k.slice(0, i)), codigo = k.slice(i + 1);
    if (!l) return;
    const ln = partesLote(l).pendientes.find((x) => x.codigo === codigo);
    if (!ln) return;
    const comp = document.getElementById('mc-' + v)?.value.trim();
    if (!comp) { avisar('Escriba el N° de comprobante del pago manual', 'error'); return; }
    if (!confirm(`¿Aplicar el pago manual a ${ln.prov.nombre} por ${money(ln.total, l.moneda)} con el comprobante ${comp}?`)) return;
    const r = S.aplicarLote({ ...l, lineas: [ln], total: ln.total }, { comprobante: comp, fecha: hoyISO() });
    if (!resultado(r, 'Pago manual aplicado')) return;
    delete ui.forma[k];
    render();
    adjuntarDespues('man|' + k, `${limpioArchivo(ln.prov.nombre).slice(0, 40)} ${comp}`, (a) => S.guardarDocLinea(r.loteId, codigo, { archivo: a }));
  },
  'ap-toggle': (v) => { ui.apAbierto[v] = !ui.apAbierto[v]; render(); },
  'ap-aviso': async (v) => {
    const i = v.indexOf('|'), loteId = v.slice(0, i), codigo = v.slice(i + 1);
    abrirDialogo(`<header class="dlg-cab"><h2>Aviso de pago</h2><button class="btn-cerrar" data-act="cerrar" aria-label="Cerrar">×</button></header>
      <div class="dlg-cuerpo"><p class="vacio">Preparando el aviso de pago…</p></div>`);
    try { mostrarArchivo(await avisoPago(loteId, codigo), 'Aviso de pago', 'Revíselo y compártalo con el proveedor.'); }
    catch (err) { cerrarDialogo(); avisar('No se pudo crear el aviso: ' + err.message, 'error'); }
  },
  'ver-arch': (v) => verComprobante(v),
  'arch-compartir': async () => {
    const a = ui.archivoAbierto;
    if (!a) return;
    try { await navigator.share({ files: [a], title: a.name.replace(/\.[a-z]+$/i, '') }); }
    catch (err) { if (err.name !== 'AbortError') { descargar(a.name, a); avisar('No se pudo compartir; se descargó el archivo'); } }
  },
  'arch-descargar': () => { if (ui.archivoAbierto) descargar(ui.archivoAbierto.name, ui.archivoAbierto); },
  'ir-prov': (v) => { location.hash = '#/proveedores'; setTimeout(() => dialogoProveedor(v), 0); },

  'f-soc': (v) => { ui.fact.soc = v; ui.fact.limite = 150; render(); },
  'f-mas': () => { ui.fact.limite += 300; render(); },
  'f-abrir': (v) => dialogoFactura(v),
  'r-mon': (v) => { ui.rep.mon = v; render(); },
  'r-prov-abrir': (k) => {
    ui.provDrop = ui.provDrop === k ? null : k; render();
    if (ui.provDrop) setTimeout(() => document.querySelector('[data-in="r-provfiltro"]')?.focus(), 0);
  },
  'r-prov-set': (v, el) => { (el.dataset.k === 'ec' ? ui.ec : ui.rep).prov = v; ui.provDrop = null; render(); },
  'r-limpiar': () => { Object.assign(ui.rep, { soc: '', prov: '', desde: '', hasta: '' }); ui.provDrop = null; render(); },
  'r-preset': (v) => { const [d, h] = v.split('|'); ui.rep.desde = d; ui.rep.hasta = h; render(); },
  'r-tipo': (v) => { ui.rep.tipo = v; ui.provDrop = null; render(); },
  'ec-mon': (v) => { ui.ec.mon = v; render(); },
  'ec-alcance': (v) => { ui.ec.alcance = v; render(); },
  'ec-limpiar': () => { Object.assign(ui.ec, { prov: '', soc: '', desde: '', hasta: '', alcance: 'saldo' }); ui.provDrop = null; render(); },
  'ec-excel': () => { try { excelEstado(); } catch (err) { avisar('No se pudo crear el Excel: ' + err.message, 'error'); } },
  'ec-pdf': () => { try { const f = pdfEstado(); descargar(f.name, f); } catch (err) { avisar('No se pudo crear el PDF: ' + err.message, 'error'); } },
  'ec-compartir': () => compartirArchivo(pdfEstado),
  'r-excel-lista': () => { try { excelListaReporte(); } catch (err) { avisar('No se pudo crear el Excel: ' + err.message, 'error'); } },
  'r-imprimir': (v) => {
    const antes = document.title;
    document.title = v === 'ec' ? nombreArchivoEstado() : nombreArchivoReporte(); // nombre sugerido al guardar como PDF
    window.print();
    setTimeout(() => { document.title = antes; }, 1000);
  },
  'r-pdf': () => { try { const f = pdfReporte(); descargar(f.name, f); } catch (err) { avisar('No se pudo crear el PDF: ' + err.message, 'error'); } },
  'r-compartir': () => compartirArchivo(pdfReporte),
  'r-excel': () => { try { excelReporte(); } catch (err) { avisar('No se pudo crear el Excel: ' + err.message, 'error'); } },
  'pn-abrir': (v) => dialogoPagoNorman(v),
  'f-nueva': () => dialogoNuevaFactura(),
  'imp-confirmar': () => confirmarImportacion(),
  'ir-sociedad': (v) => { ui.fact.soc = v; ui.fact.est = 'saldo'; ui.fact.limite = 150; location.hash = '#/facturas'; },

  'pv-filtro': (v) => { ui.prov.filtro = v; render(); },
  'pv-abrir': (v) => dialogoProveedor(v),
  'pv-nuevo': () => dialogoProveedor(null),

  'cfg-excel': () => exportarExcel(),
};

let tBusqueda;
/** Puntos de miles mientras se escribe, sin mover el cursor. Un "." tecleado se toma como coma de céntimos. */
function formatearCampoMonto(el, e) {
  let v = el.value, pos = el.selectionStart ?? v.length;
  if (e.data === '.' && pos > 0 && v[pos - 1] === '.') v = v.slice(0, pos - 1) + ',' + v.slice(pos);
  const antes = v.slice(0, pos).replace(/[^\d,]/g, '').length; // dígitos y coma antes del cursor
  const nuevo = formatearMontoEscrito(v);
  if (nuevo === el.value) return;
  el.value = nuevo;
  let i = 0, cuenta = 0;
  while (i < nuevo.length && cuenta < antes) { if (/[\d,]/.test(nuevo[i])) cuenta++; i++; }
  try { el.setSelectionRange(i, i); } catch { /* sin cursor */ }
}

// Al tocar un monto se selecciona todo (para escribir encima); al salir se completan los céntimos.
document.addEventListener('focusin', (e) => {
  const el = e.target;
  if (!el.matches?.('[data-monto]')) return;
  try { el.select(); } catch { /* */ }
  el.dataset.seleccionar = '1'; // el clic que dio el foco no debe quitar la selección
});
document.addEventListener('mouseup', (e) => {
  const el = e.target;
  if (el.matches?.('[data-monto]') && el.dataset.seleccionar) { e.preventDefault(); delete el.dataset.seleccionar; try { el.select(); } catch { /* */ } }
});
document.addEventListener('touchend', (e) => {
  const el = e.target;
  if (el.matches?.('[data-monto]') && el.dataset.seleccionar) { delete el.dataset.seleccionar; setTimeout(() => { try { el.setSelectionRange(0, el.value.length); } catch { /* */ } }, 0); }
});
document.addEventListener('focusout', (e) => {
  const el = e.target;
  if (!el.matches?.('[data-monto]') || !el.value.trim()) return;
  const n = parseMonto(el.value);
  if (!(n >= 0)) return;
  const v = fmtNum(n);
  if (v !== el.value) { el.value = v; el.dispatchEvent(new Event('input', { bubbles: true })); }
});

document.addEventListener('input', (e) => {
  const el = e.target, k = el.dataset.in;
  if (el.matches?.('[data-monto]')) formatearCampoMonto(el, e);
  if (!k) return;
  if (k === 'p-monto') { ui.prog.sel.set(el.dataset.v, el.value); actualizarBarra(); return; }
  if (k === 'iban') { validarCampoIBAN(el); return; }
  if (k === 'r-provfiltro') {
    const t = el.value.trim().toLowerCase();
    let n = 0;
    document.querySelectorAll('.rep-prov-lista li[data-nombre]').forEach((li) => { const ok = !t || li.dataset.nombre.includes(t); li.hidden = !ok; if (ok) n++; });
    const nada = document.querySelector('.rep-prov-nada'); if (nada) nada.hidden = n > 0;
    return;
  }
  if (k === 'fm-monto') { actualizarRestoMov(); return; }
  if (k === 'pn-monto') { const r = document.getElementById('pn-resto'); if (r) r.innerHTML = textoResto(S.factura(el.dataset.v), parseMonto(el.value)); return; }
  if (k === 'nf-fecha' || k === 'nf-plazo') {
    const form = el.form;
    const ff = form.fechaFactura.value, pl = Number(form.plazo.value) || 0;
    if (ff) form.fechaVence.value = sumarDias(ff, pl);
    return;
  }
  const busquedas = { 'c-q': (x) => { ui.canc.q = x; ui.canc.limite = 100; }, 'p-q': (x) => { ui.prog.q = x; }, 'f-q': (x) => { ui.fact.q = x; ui.fact.limite = 150; }, 'pv-q': (x) => { ui.prov.q = x; } };
  if (busquedas[k]) {
    busquedas[k](el.value);
    clearTimeout(tBusqueda);
    tBusqueda = setTimeout(render, 300);
  }
});

document.addEventListener('change', async (e) => {
  const el = e.target, k = el.dataset.in;
  if (!k) return;
  if (k === 'fm-tipo') { actualizarRestoMov(); return; }
  if (k === 'pn-forma') { el.form.dataset.forma = el.value; return; }
  if (k === 'adj-file') {
    const file = el.files[0], lbl = el.closest('.adjuntar');
    if (!file) return;
    if (file.size > 8 * 1024 * 1024) { avisar('El archivo pesa más de 8 MB', 'error'); el.value = ''; return; }
    ui.adj[el.dataset.k] = file;
    if (lbl) { lbl.classList.add('listo'); lbl.querySelector('span').innerHTML = `<b>${esc(file.name)}</b> · ${Math.max(1, Math.round(file.size / 1024))} KB`; }
    return;
  }
  if (k === 'ap-doc') {
    const i = el.dataset.v.indexOf('|');
    resultado(S.guardarDocLinea(el.dataset.v.slice(0, i), el.dataset.v.slice(i + 1), { documento: el.value.trim() }), 'N° de documento guardado');
    return;
  }
  if (k === 'ap-pdf' || k === 'ap-gen') {
    const file = el.files[0];
    el.value = '';
    if (!file) return;
    const v = el.dataset.v, i = v.indexOf('|');
    const loteId = k === 'ap-gen' ? v : v.slice(0, i), codigo = v.slice(i + 1);
    const l = S.datos().lotes[loteId];
    avisar('Subiendo comprobante…');
    try {
      const nombre = k === 'ap-gen' ? `Debito ${l?.comprobante || ''}` : `${limpioArchivo(S.proveedor(codigo).nombre).slice(0, 40)} ${l?.comprobante || ''}`;
      const a = await subirComprobante(file, nombre);
      const r = k === 'ap-gen' ? S.guardarAdjLote(loteId, a) : S.guardarDocLinea(loteId, codigo, { archivo: a });
      if (resultado(r, 'Comprobante adjuntado')) render();
    } catch (err) { avisar('No se pudo adjuntar: ' + err.message, 'error'); }
    return;
  }
  if (k === 'r-soc') { ui.rep.soc = el.value; render(); return; }
  if (k === 'r-desde') { ui.rep.desde = el.value; render(); return; }
  if (k === 'ec-soc') { ui.ec.soc = el.value; render(); return; }
  if (k === 'ec-desde') { ui.ec.desde = el.value; render(); return; }
  if (k === 'ec-hasta') { ui.ec.hasta = el.value; render(); return; }
  if (k === 'r-hasta') { ui.rep.hasta = el.value; render(); return; }
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
  if (k === 'c-mon') { ui.canc.mon = el.value; ui.canc.limite = 100; render(); return; }
  if (k === 'imp-col') { ui.imp.mapeo[el.dataset.v] = el.value === '' ? '' : Number(el.value); dialogoImportar(); return; }
  if (k === 'imp-prov') { ui.imp.asignados[el.dataset.v] = el.value; dialogoImportar(); return; }
  if (k === 'imp-fijo') { ui.imp.fijos[el.dataset.v] = el.value; dialogoImportar(); return; }
  if (k === 'imp-archivo') {
    const file = el.files[0];
    el.value = '';
    if (!file) return;
    try { await cargarArchivoImportar(file); } catch (err) { avisar('No se pudo leer el archivo: ' + err.message, 'error'); }
    return;
  }
  if (k === 'a-extra' || k === 'a-desc') {
    const clave = decodeURIComponent(el.dataset.lote);
    if (k === 'a-extra') {
      const n = el.value.trim() ? parseMonto(el.value) : 0;
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
    Sync.configurarConexion(d.url || API_URL, d.servicio).then(() => entrarComo(d.usuario)).catch((err) => {
      avisar(navigator.onLine ? err.message : 'La primera vez en este dispositivo se necesita internet', 'error');
      btn.disabled = false;
      btn.textContent = 'Entrar';
    });
  } else if (k === 'pedir-revision') {
    const id = form.dataset.v, nota = d.nota.trim();
    const f = S.factura(id);
    if (!S.programacionDe(f)) {
      const monto = parseMonto(d.monto);
      if (!(monto > 0) || monto > S.saldo(f) + 0.004) { avisar('Revise el monto: debe ser mayor que cero y no más que el saldo', 'error'); return; }
      const r = S.programarPagos([{ facturaId: id, monto }], d.fecha);
      if (r.error) { avisar(r.error, 'error'); return; }
      ui.prog.sel.delete(id);
    }
    if (resultado(S.actualizarFactura(id, { revision: true, observaciones: nota }), 'Revisión pedida: no se paga hasta que la revisen')) { cerrarDialogo(); render(); }
  } else if (k === 'devolver-revision') {
    const id = form.dataset.v, motivo = d.motivo.trim();
    const p = S.programacionDe(S.factura(id));
    if (p) { const r = S.quitarProgramacion(p.id, 'Revisión: ' + motivo); if (r.error) { avisar(r.error, 'error'); return; } }
    if (resultado(S.actualizarFactura(id, { revision: false, observaciones: 'Revisión: ' + motivo }), 'Devuelta a Norman')) { cerrarDialogo(); render(); }
  } else if (k === 'u-crear') {
    Sync.adminUsuarios({ op: 'crear', usuario: d.usuario.trim().toLowerCase(), nombre: d.nombre.trim(), rol: d.rol })
      .then((r) => { form.reset(); mostrarPin(d.nombre.trim(), r.pin); })
      .catch((e) => avisar(navigator.onLine ? e.message : 'Para crear usuarios se necesita internet', 'error'));
  } else if (k === 'ver-conexion') {
    if (Sync.verificarClaveServicio(d.clave)) dialogoConexion(true);
    else avisar('La contraseña del servicio no es correcta', 'error');
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
    if (d.nueva !== d.repetir) { avisar('Los dos PIN nuevos no coinciden', 'error'); return; }
    if (!/^\d{4}$/.test(d.nueva)) { avisar('El PIN debe tener exactamente 4 números', 'error'); return; }
    if (/^(\d)\1{3}$/.test(d.nueva) || '0123456789'.includes(d.nueva) || '9876543210'.includes(d.nueva)) {
      avisar('Elija un PIN menos fácil de adivinar (no 1111, 1234, 4321…)', 'error'); return;
    }
    const btn = form.querySelector('button[type=submit]');
    btn.disabled = true;
    Sync.cambiarContrasena(d.actual || '', d.nueva).then(() => {
      cerrarDialogo();
      avisar('PIN guardado');
      render();
    }).catch((err) => { avisar(navigator.onLine ? err.message : 'Para cambiar el PIN se necesita internet', 'error'); btn.disabled = false; });
  } else if (k === 'f-fechas') {
    const f = S.factura(form.dataset.v);
    const cambios = { fechaFactura: d.fechaFactura || '', observaciones: d.observaciones, fechaVence: d.fechaVence || '' };
    if (resultado(S.actualizarFactura(f.id, cambios), 'Cambios guardados')) dialogoFactura(f.id);
  } else if (k === 'pago-norman') {
    const f = S.factura(form.dataset.v);
    const x = valoresPagoNorman(form);
    if (x.forma === 'banco') {
      const l = S.loteDirecto(f, x.monto, x.fecha, x.descripcion);
      const prob = S.problemasLote(l);
      if (prob.length) { avisar(prob.join(' '), 'error'); return; }
      const n = leerConsecutivo(form.consecutivo);
      if (!n) return;
      try { textoArchivo(l, n); } catch (err) { avisar(err.message, 'error'); return; }
      const r = S.pagarAhoraBanco(f, { ...x, consecutivo: n });
      if (r.error) { avisar(r.error, 'error'); return; }
      try { descargarTxt(r.lote, n, '_FACT' + String(f.documento).replace(/[^A-Za-z0-9]+/g, '')); } catch (err) { avisar(err.message, 'error'); }
      avisar(`Archivo descargado (consecutivo ${c3(n)}) · el pago quedó en «Por aplicar» esperando el comprobante del banco`);
      cerrarDialogo(); render();
    } else {
      if (!String(x.comprobante || '').trim()) { avisar('Escriba el N° de comprobante', 'error'); return; }
      const r = S.pagarAutorizadoNorman(f, x);
      if (!resultado(r, `${r.tipo === 'abono' ? 'Abono' : 'Pago'} aplicado`)) return;
      cerrarDialogo(); render();
      adjuntarDespues('pn|' + f.id, `${limpioArchivo(S.proveedor(f.codigo).nombre).slice(0, 40)} ${x.comprobante}`, (a) => S.guardarDocLinea(r.loteId, r.codigo, { archivo: a }));
    }
  } else if (k === 'f-mov') {
    if (resultado(S.registrarMovimiento(form.dataset.v, { tipo: d.tipo, monto: parseMonto(d.monto), fecha: d.fecha, comprobante: d.comprobante }), 'Movimiento registrado')) dialogoFactura(form.dataset.v);
  } else if (k === 'f-crear') {
    const codigo = (d.proveedor.match(/·\s*(\S+)$/) || [])[1];
    if (!codigo || !S.datos().proveedores[codigo]) { avisar('Elija el proveedor de la lista (o créelo en Proveedores)', 'error'); return; }
    const monto = parseMonto(d.monto);
    if (!(monto > 0)) { avisar('Monto inválido', 'error'); return; }
    const r = S.nuevaFactura({ sociedad: d.sociedad, codigo, documento: d.documento.trim(), moneda: d.moneda, fechaFactura: d.fechaFactura, fechaVence: d.fechaVence || sumarDias(d.fechaFactura, 30), monto, observaciones: d.observaciones });
    if (resultado(r, 'Factura guardada')) cerrarDialogo();
  } else if (k === 'pv-guardar') {
    const nuevo = !form.querySelector('[name=codigo]').readOnly;
    if (resultado(S.guardarProveedor(d.codigo.trim(), { nombre: d.nombre, alias: d.alias, ibanCRC: d.ibanCRC, ibanUSD: d.ibanUSD, plazo: d.plazo }, nuevo), 'Proveedor guardado')) cerrarDialogo();
  } else if (k === 'cfg-banco') {
    const ced = d.controlEncabezado.replace(/\D/g, '');
    if (d.controlEncabezado.trim() && (!ced || ced.length > 12)) { avisar('La cédula del empleado debe tener solo números (hasta 12)', 'error'); return; }
    resultado(S.guardarConfig({ controlEncabezado: ced }), 'Configuración guardada');
  } else if (k === 'cfg-soc') {
    const soc = form.dataset.v;
    const cedula = d.cedulaBCR.replace(/\D/g, '');
    if (d.cedulaBCR.trim() && (!cedula || cedula.length > 15)) { avisar('La cédula jurídica debe tener solo números', 'error'); return; }
    const ultimo = d.ultimoConsec.trim() === '' ? null : Number(d.ultimoConsec.trim());
    if (ultimo != null && !(Number.isInteger(ultimo) && ultimo >= 0 && ultimo <= 999)) { avisar('El consecutivo debe ser un número de 0 a 999', 'error'); return; }
    const r = S.guardarSociedad(soc, d);
    if (r.error) { avisar(r.error, 'error'); return; }
    const b = S.bancoSoc(soc);
    if (b.cedula !== cedula || (b.ultimo ?? null) !== ultimo) { const r2 = S.guardarBancoSoc(soc, { cedula, ultimo }); if (r2.error) { avisar(r2.error, 'error'); return; } }
    avisar('Sociedad guardada'); render();
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

/** Entra como la persona elegida y abre el Tablero. */
let entrando = false;
async function entrarComo(usuario) {
  if (entrando) return;
  entrando = true;
  document.querySelectorAll('.tarjeta-usuario').forEach((b) => { b.disabled = true; });
  const t = document.querySelector(`[data-act=elegir][data-v="${CSS.escape(usuario)}"] small`);
  if (t) t.textContent = 'Entrando…';
  try {
    await Sync.entrarDirecto(usuario);
    ui.login = '';
    location.hash = '#/tablero';
    render();
  } catch (e) {
    avisar(e.message, 'error');
    ui.login = '';
    render();
  } finally {
    entrando = false;
  }
}

/** Muestra el resultado de un cambio. Devuelve true si se aplicó. */
function resultado(r, okMsg) {
  if (r.error) { avisar(r.error, 'error'); return false; }
  avisar(okMsg + (r.avisos?.length ? ' · ' + r.avisos.join(' · ') : ''));
  return true;
}

document.getElementById('pie').textContent = `Cuentas por Pagar · Grupo BSE · Versión ${VERSION}`;

// ---------------------------------------------------------------- Aviso de versión nueva

let avisoVersion = false;
async function revisarVersion() {
  if (avisoVersion || !navigator.onLine) return;
  try {
    const r = await fetch('version.json?t=' + Date.now(), { cache: 'no-store' });
    if (!r.ok) return;
    const { version } = await r.json();
    if (version && version !== VERSION) mostrarAvisoVersion(version);
  } catch { /* sin conexión */ }
}

function mostrarAvisoVersion(version) {
  avisoVersion = true;
  const b = document.createElement('div');
  b.className = 'aviso-version';
  b.setAttribute('role', 'status');
  b.innerHTML = `<span>Hay una versión nueva de la aplicación (${esc(version)}).</span><button class="btn primario chico" data-act="actualizar">Actualizar</button>`;
  document.body.prepend(b);
}

async function actualizarApp() {
  if (Sync.estado.pendientes.length && navigator.onLine) await Sync.sincronizar();
  try {
    const ks = await caches.keys();
    await Promise.all(ks.map((k) => caches.delete(k)));
    const reg = await navigator.serviceWorker?.getRegistration();
    await reg?.update();
  } catch { /* sin caché */ }
  location.reload();
}

setTimeout(revisarVersion, 4000);
setInterval(revisarVersion, 5 * 60 * 1000);
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') revisarVersion(); });

Sync.alAvisar((msg, tipo) => avisar(msg, tipo === 'aviso' ? 'ok' : tipo));
Sync.iniciar();
render();

if ('serviceWorker' in navigator && location.protocol === 'https:') {
  navigator.serviceWorker.register('sw.js').catch((e) => console.warn('Sin modo sin internet:', e));
}
