// Pantallas y eventos de la aplicación.
import * as S from './store.js';
import * as Sync from './sync.js';
import { VERSION, USUARIOS, API_URL } from './config.js';
import { importarExcel } from './importer.js';
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
  rep: { soc: '', prov: '', mon: 'CRC', desde: '', hasta: '' },
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
    parchar($app, t.content);
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
function selectorDe(el) {
  if (el.id) return '#' + CSS.escape(el.id);
  if (el.dataset.in) return `[data-in="${el.dataset.in}"]` + (el.dataset.v ? `[data-v="${CSS.escape(el.dataset.v)}"]` : '');
  if (el.name) return `${el.tagName.toLowerCase()}[name="${CSS.escape(el.name)}"]`;
  return null;
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
        <td class="num"><input class="mini" type="text" inputmode="decimal" data-monto value="${ln.extra ? fmtNum(ln.extra) : ''}" placeholder="0,00" data-in="a-extra" data-lote="${id}" data-v="${esc(ln.codigo)}" aria-label="Intereses u otros cargos para ${esc(ln.prov.nombre)}"></td>
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
      ${s > 0.004 && !S.esNorman() ? `<div class="acciones-fila"><button class="btn btn-norman" data-act="pn-abrir" data-v="${f.id}">✔ Autorizado por Norman</button></div>` : ''}
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
          const tagNorman = S.esAutorizadoNorman(prog) ? ' <span class="tag-norman">Autorizado por Norman</span>' : '';
          const tipoTxt = String(prog?.nota || '').startsWith(S.NOTA_NC) ? 'Nota de crédito' : (tipos[m.tipo] || m.tipo);
          return `<tr><td>${fmtFecha(m.fecha) || '<span class="sub">sin fecha</span>'}</td><td>${tipoTxt}${tagNorman}</td><td>${esc(m.comprobante)}</td>
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
      <div><h2>Autorizado por Norman</h2>
        <p class="sub">Factura ${esc(f.documento)} · ${esc(prov.nombre)} · ${esc(S.nombreSociedad(f.sociedad))} · saldo ${money(s, f.moneda)}</p></div>
      <button class="btn-cerrar" data-act="cerrar" aria-label="Cerrar">×</button></header>`;
  if (ya) {
    abrirDialogo(`${cab}<div class="dlg-cuerpo">
      <p class="aviso-pruebas">Esta factura ya está programada por ${esc(ya.por)} por ${money(ya.monto, f.moneda)} para el ${fmtFecha(ya.fechaPago)}. Aplíquela desde «Por aplicar» para no pagarla dos veces.</p>
      <div class="acciones-fila"><a class="btn primario" href="#/aplicar" data-act="cerrar">Ir a Por aplicar</a></div></div>`);
    return;
  }
  const lote = S.loteDirecto(f, s, hoyISO());
  abrirDialogo(`${cab}
    <form class="dlg-cuerpo form-grid" data-form="pago-norman" data-v="${f.id}">
      <label>Monto a pagar <input type="text" name="monto" inputmode="decimal" data-monto value="${fmtNum(s)}" required data-in="pn-monto" data-v="${f.id}">
        <span class="pn-resto" id="pn-resto">${textoResto(f, s)}</span></label>
      <label>Fecha de pago <input type="date" name="fecha" value="${hoyISO()}" required></label>
      <label class="ancho">Nota <span class="sub-inline">opcional</span> <input type="text" name="nota" placeholder="Ej. pagar hoy, abono acordado"></label>
      <div class="ancho pn-banco">
        <strong>Archivo para el banco</strong>
        <p class="sub">Si paga con archivo, descárguelo y súbalo al banco. Si hace la transferencia a mano, no lo necesita.</p>
        <label>Descripción en el banco <input type="text" name="descripcion" maxlength="60" value="${esc(lote.lineas[0].descripcion)}"></label>
        ${S.problemasLote(lote).length ? `<div class="aviso"><strong>Para generar el archivo falta:</strong><ul>${S.problemasLote(lote).map((x) => `<li>${esc(x)}</li>`).join('')}</ul></div>` : ''}
        <button class="btn oscuro" type="button" data-act="pn-txt" data-v="${f.id}" ${S.problemasLote(lote).length ? 'disabled' : ''}>Descargar .txt de este pago</button>
      </div>
      <label class="ancho">N° de comprobante del banco <input type="text" name="comprobante" required autocomplete="off"></label>
      <div class="ancho acciones-fila"><button class="btn primario" type="submit">Aplicar pago</button>
        <button class="btn" type="button" data-act="f-abrir" data-v="${f.id}">Volver</button></div>
      <p class="ancho sub">Queda registrado como «Autorizado por Norman», con su nombre como quien aplicó. Norman lo ve en el Tablero.</p>
    </form>`);
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
  return { monto: parseMonto(d.monto), fecha: d.fecha, nota: d.nota, comprobante: d.comprobante, descripcion: d.descripcion };
}

function seccionAutorizados() {
  const lista = S.autorizadosNorman(30);
  if (!lista.length) return '';
  return `<section class="bloque bloque-norman">
    <h2>Autorizados por Norman <span class="sub-inline">últimos 30 días</span></h2>
    <p class="sub">Pagos que Norman indicó de palabra y se pagaron de inmediato.</p>
    <div class="tabla-scroll"><table class="tabla">
      <thead><tr><th>Fecha</th><th>Sociedad</th><th>Proveedor</th><th>Factura</th><th class="num">Monto</th><th>Aplicó</th><th>Comprobante</th><th>Nota</th></tr></thead>
      <tbody>${lista.map(({ p, f }) => `<tr data-act="f-abrir" data-v="${f.id}" tabindex="0">
        <td>${fmtFecha(p.fechaPago)}</td><td>${esc(S.nombreSociedad(f.sociedad))}</td><td>${esc(S.proveedor(f.codigo).nombre)}</td>
        <td>${esc(f.documento)}</td><td class="num">${money(p.monto, f.moneda)}</td><td>${esc(p.cerradoPor)}</td>
        <td>${esc(p.comprobante)}</td><td>${esc(p.nota.slice(S.NOTA_NORMAN.length).replace(/^\s*·\s*/, ''))}</td></tr>`).join('')}</tbody>
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

function vistaReportes() {
  const R = ui.rep;
  const socs = Object.keys(S.sociedades());
  const grupos = datosReporte();
  const emitido = new Date();
  const mon = MONEDAS[R.mon];
  const hoy = hoyISO();
  const preset = (n, d, h) => `<button class="chip ${R.desde === d && R.hasta === h ? 'on' : ''}" data-act="r-preset" data-v="${d}|${h}">${n}</button>`;
  const total = { n: grupos.reduce((t, g) => t + g.cantidad, 0), monto: round2(grupos.reduce((t, g) => t + g.monto, 0)), saldo: round2(grupos.reduce((t, g) => t + g.saldo, 0)) };
  return `
  <div class="titulo-vista no-imprimir"><div><h1>Reportes</h1><p class="sub">Cuentas por pagar con saldo, por fecha de vencimiento. Un reporte por moneda.</p></div>
    <div class="acciones"><button class="btn" data-act="r-excel" ${total.n ? '' : 'disabled'}>Descargar Excel</button>
      <button class="btn primario" data-act="r-imprimir" ${total.n ? '' : 'disabled'}>Imprimir o guardar PDF</button></div></div>
  <section class="filtros no-imprimir">
    <div class="form-grid rep-filtros">
      <label>Sociedad <select data-in="r-soc"><option value="">Todas las sociedades</option>${socs.map((x) => `<option value="${esc(x)}" ${R.soc === x ? 'selected' : ''}>${esc(S.nombreSociedad(x))}</option>`).join('')}</select></label>
      <div class="rep-prov-drop"><span class="lbl">Proveedor</span>
        <button type="button" class="rep-prov-boton ${R.prov ? 'elegido' : ''}" data-act="r-prov-abrir" aria-expanded="${ui.rep.provAbierto ? 'true' : 'false'}">
          <span>${esc(R.prov ? S.proveedor(R.prov).nombre : 'Todos los proveedores')}</span><span aria-hidden="true">▾</span></button>
        ${ui.rep.provAbierto ? `<div class="rep-prov-panel">
          <input type="search" data-in="r-provfiltro" placeholder="Escriba para buscar…" autocomplete="off" autocorrect="off" autocapitalize="characters" spellcheck="false">
          <ul class="rep-prov-lista">
            <li><button type="button" data-act="r-prov-set" data-v="" class="${R.prov ? '' : 'on'}">(Todos los proveedores)</button></li>
            ${proveedoresReporte().map((p) => `<li data-nombre="${esc(p.nombre.toLowerCase())}"><button type="button" data-act="r-prov-set" data-v="${esc(p.codigo)}" class="${R.prov === p.codigo ? 'on' : ''}">${esc(p.nombre)}</button></li>`).join('')}
          </ul>
          <p class="sub rep-prov-nada" hidden>Ningún proveedor coincide.</p>
        </div>` : ''}
      </div>
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
          <thead><tr><th class="w-doc">N°<span class="larga"> Factura</span></th><th class="c-ff w-ff">Fecha de<br>factura</th><th class="w-fv"><span class="larga">Fecha de<br>vencimiento</span><span class="corta">Vence</span></th><th class="num w-dias">Días</th><th class="num w-saldo">Saldo</th><th class="c-mon w-mon">Moneda</th></tr></thead>
          <tbody>${x.facturas.map((f) => `<tr><td>${esc(f.documento)}</td><td class="c-ff">${fmtFecha(f.fechaFactura)}</td><td><span class="larga">${fmtFecha(f.fechaVence)}</span><span class="corta">${fmtFecha(f.fechaVence).replace(/\/20(\d\d)$/, '/$1')}</span></td><td class="num">${celdaDias(f)}</td>
            <td class="num">${fmtNum(S.saldo(f))}</td><td class="c-mon">${mon.nombre}</td></tr>`).join('')}</tbody>
          ${R.prov ? '' : `<tfoot><tr class="rep-tot-num"><td></td><td class="c-ff"></td><td></td><td class="rep-tot-junto"><span>Total Cliente · ${x.facturas.length} ${x.facturas.length === 1 ? 'factura' : 'facturas'}</span></td><td class="num">${fmtNum(x.saldo)}</td><td class="c-mon"></td></tr></tfoot>`}
        </table>
      </section>`).join('')}
      <table class="rep-tabla rep-tabla-total rep-total-soc"><tbody><tr><td class="w-doc"></td><td class="c-ff w-ff"></td><td class="w-fv"></td><td class="w-dias rep-tot-junto"><span>Total ${esc(S.nombreSociedad(g.soc))} · ${g.cantidad} ${g.cantidad === 1 ? 'factura' : 'facturas'}</span></td><td class="num w-saldo">${fmtNum(g.saldo)}</td><td class="c-mon w-mon"></td></tr></tbody></table>
    </article>`).join('')}</div>
  ${!R.soc && grupos.filter((g) => g.cantidad).length > 1 ? `<div class="reporte rep-general"><table class="rep-tabla rep-tabla-total rep-total-soc"><tbody><tr><td class="w-doc"></td><td class="c-ff w-ff"></td><td class="w-fv"></td><td class="w-dias rep-tot-junto"><span>Total todas las sociedades · ${total.n} facturas</span></td><td class="num w-saldo">${fmtNum(total.saldo)}</td><td class="c-mon w-mon"></td></tr></tbody></table></div>` : ''}`;
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
      if (celda && celda.t === 'n') celda.z = c === 3 ? '+0;-0;"Hoy"' : '#,##0.00';
    }
  }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Reporte');
  const soc = (R.soc ? S.nombreSociedad(R.soc).replace(/[^A-Za-z0-9]+/g, '') : 'Todas') + (R.prov ? '_' + S.proveedor(R.prov).nombre.replace(/[^A-Za-z0-9]+/g, '').slice(0, 30) : '');
  const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  descargar(`Reporte_CxP_${soc}_${R.mon}_${hoy}.xlsx`, new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
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
  if (ui.rep.provAbierto && !e.target.closest('.rep-prov-drop')) { ui.rep.provAbierto = false; render(); }
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
  if (e.key === 'Escape' && ui.rep.provAbierto) { ui.rep.provAbierto = false; render(); return; }
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
  'r-mon': (v) => { ui.rep.mon = v; render(); },
  'r-prov-abrir': () => {
    ui.rep.provAbierto = !ui.rep.provAbierto; render();
    if (ui.rep.provAbierto) setTimeout(() => document.querySelector('[data-in="r-provfiltro"]')?.focus(), 0);
  },
  'r-prov-set': (v) => { ui.rep.prov = v; ui.rep.provAbierto = false; render(); },
  'r-limpiar': () => { Object.assign(ui.rep, { soc: '', prov: '', desde: '', hasta: '', provAbierto: false }); render(); },
  'r-preset': (v) => { const [d, h] = v.split('|'); ui.rep.desde = d; ui.rep.hasta = h; render(); },
  'r-imprimir': () => window.print(),
  'r-excel': () => { try { excelReporte(); } catch (err) { avisar('No se pudo crear el Excel: ' + err.message, 'error'); } },
  'pn-abrir': (v) => dialogoPagoNorman(v),
  'pn-txt': (v, el) => {
    const f = S.factura(v);
    const x = valoresPagoNorman(el.closest('form'));
    if (!(x.monto > 0)) { avisar('Monto inválido', 'error'); return; }
    try {
      const l = S.loteDirecto(f, x.monto, x.fecha, x.descripcion);
      const prob = S.problemasLote(l);
      if (prob.length) { avisar(prob.join(' '), 'error'); return; }
      const soc = S.nombreSociedad(f.sociedad).toUpperCase().replace(/[^A-Z0-9]+/g, '');
      descargar(`PROVEEDORES_${soc}_${f.moneda}_${x.fecha.split('-').reverse().join('')}_FACT${String(f.documento).replace(/[^A-Za-z0-9]+/g, '')}.txt`, textoArchivo(l), 'text/plain');
    } catch (err) { avisar(err.message, 'error'); }
  },
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
  if (el.matches?.('[data-monto]')) setTimeout(() => { try { el.select(); } catch { /* */ } }, 0);
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
  if (k === 'r-soc') { ui.rep.soc = el.value; render(); return; }
  if (k === 'r-desde') { ui.rep.desde = el.value; render(); return; }
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
    const r = S.pagarAutorizadoNorman(f, valoresPagoNorman(form));
    if (resultado(r, `${r.tipo === 'abono' ? 'Abono' : 'Pago'} aplicado · Autorizado por Norman`)) { cerrarDialogo(); render(); }
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
