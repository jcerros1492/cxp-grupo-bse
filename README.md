# Cuentas por Pagar · Grupo BSE

Aplicación web para las cuentas por pagar de Santa Eduviges, La Hilda, Nosavar y Loma Bonita.

- La **base de datos es una hoja de Google Sheets**, que se puede abrir y consultar en cualquier momento.
- **Funciona sin internet.** Cada computadora guarda una copia y los cambios se envían a la hoja cuando vuelve la conexión.
- **Sincroniza entre las tres personas.** Cada 20 segundos, y en cada cambio, la aplicación trae lo que hicieron los demás.

## Versiones

La versión publicada se ve al pie de cada pantalla. Si no coincide con la última, falta subir los archivos a GitHub o recargar con Ctrl + F5.

| Versión | Cambios |
|---|---|
| 1.18.0 | **Banco o Manual.** En «Por aplicar» cada proveedor se paga por Banco (.txt) o Manual. Al descargar el .txt el pago queda «Enviado al banco · esperando comprobante» (quién y cuándo, visible en todos los dispositivos), con aviso si se intenta descargar de nuevo y botón «Volver a pendiente». El comprobante general del débito lo escribe Jonathan, Fernando o Daniel. «Pagar ahora» (antes «Autorizado por Norman») con Banco o Manual. Comprobantes en PDF o foto guardados en Google Drive (carpeta «CxP Comprobantes») y visibles desde cualquier dispositivo sin guardarse en él. En «Aplicados recientemente»: comprobante general (solo uso interno), N° de documento y comprobante individual por proveedor, y **Aviso de pago** en PDF para compartir por WhatsApp. **Requiere el bloque `apps-script/Adjuntos.gs` en Apps Script.** |
| 1.17.2 | Pestaña «Reportes» (antes «Cuentas por pagar»). «Limpiar filtros» del estado de cuenta también quita el proveedor. Totales con el signo de la moneda (₡ / $) en reportes y estado de cuenta: pantalla, Excel (formato de moneda) y PDF (¢ / $, porque la letra del PDF no trae ₡). |
| 1.17.1 | Estado de cuenta: cada abono, pago o nota de crédito debajo de su factura con «↳» y a qué factura se aplicó («Nota de crédito N° NC-56 · aplicada a la factura 2752»); línea entre facturas; columna «Monto» (antes «Cargos»); títulos centrados. En el Excel, columna «Factura» para filtrar. Reportes: títulos y cifras centrados bajo su columna (Días y Saldo incluidos) y alineados por los decimales, también en el PDF. |
| 1.17.0 | Reportes con dos pestañas: «Cuentas por pagar» y «Estado de cuenta». Estado de cuenta por proveedor (sociedad, moneda, con saldo o todas, fechas): cada factura con su monto y sus abonos, pagos y notas de crédito, saldo después de cada movimiento y totales; en pantalla, PDF (compartir/descargar), Excel e impresión. Cuentas por pagar: dos Excel a elegir, «Excel reporte» (agrupado) y «Excel lista» (una fila por factura con filtros). Nombres de archivo «Reporte CxP - PROVEEDOR - COLONES» y «Estado de cuenta - PROVEEDOR - COLONES». Al tocar un monto se selecciona completo de inmediato (antes, al escribir muy rápido, el número se pegaba al anterior). |
| 1.16.1 | Número de página («Página 1 de 3») al pie de cada hoja, en el PDF generado y al imprimir. |
| 1.16.0 | Reportes en PDF generados por la aplicación (jsPDF, sin fecha, título ni dirección del navegador): «Compartir PDF» en el celular (WhatsApp, correo…) y «Descargar PDF». Nombre del archivo con el proveedor, la sociedad, la moneda y la fecha («Reporte CxP - DON CORLEONE S.A. - Todas - Colones - 2026-10-30»), igual en el Excel. «Imprimir» sin encabezados del navegador. |
| 1.15.7 | Reportes: botón «Limpiar filtros» (sociedad, proveedor y fechas). Proveedor como filtro tipo Excel: lista desplegable con casilla para escribir y la lista se reduce al escribir. En el reporte por proveedor solo va el total de cada sociedad (sin repetir «Total Cliente»). |
| 1.15.6 | PDF: al imprimir siempre se usa el diseño de escritorio (la fecha de emisión pequeña a la derecha, también desde el celular). Reporte por proveedor: las sociedades siguen en la misma hoja, con el encabezado completo solo al inicio. Proveedor con buscador: se escribe parte del nombre y se elige de la lista. |
| 1.15.5 | Reporte: «Total Cliente · N facturas» junto al saldo de cada proveedor; total de la sociedad y general también junto al saldo; totales un poco más grandes. Igual en el Excel. |
| 1.15.4 | Reporte solo con Saldo (sin Monto), porque ya descuenta abonos y notas de crédito; también en PDF y Excel. Columna Moneda centrada. |
| 1.15.3 | Reportes: filtro por proveedor (en una sociedad o en todas); encabezados centrados, fechas en dos líneas y columnas más compactas para hoja carta; Excel con formato de miles y céntimos, días con signo y total de todas las sociedades. |
| 1.15.2 | Reporte bien alineado: todas las tablas de proveedores usan los mismos anchos de columna, y el total de la sociedad (y el general) va como fila alineada bajo Monto y Saldo. |
| 1.15.1 | Reportes: columna «Días» respecto a hoy, en rojo y negativo si está vencida y en verde y positivo si falta para vencer («Hoy» si vence hoy), con leyenda; también en el Excel como número. En el celular la tabla cabe sin deslizar (fecha corta; en pantallas muy angostas se oculta Monto y queda Saldo). |
| 1.15.0 | Menú «Reportes»: cuentas por pagar con saldo por sociedad (o todas) y por moneda, filtradas por fecha de vencimiento (desde–hasta y atajos), agrupadas por proveedor con subtotales; encabezado con la sociedad, «Reporte de Cuentas por Pagar», rango y fecha de emisión; imprimir o guardar PDF (una hoja por sociedad) y Excel. En «Registrar abono o nota de crédito» se ve en vivo cuánto saldo queda. Los montos se seleccionan completos al tocarlos y completan los céntimos al salir. |
| 1.14.0 | Notas de crédito como movimiento propio de la factura (no pasan por Lotes ni Programaciones); requiere agregar una línea en Apps Script (ver abajo). «Programar pagos» solo para Norman. Revisión: Norman pide revisar con el monto y la fecha; el pago queda retenido en «Por aplicar» hasta que Jonathan, Fernando o Daniel lo marquen «Revisado» (pasa al lote para pagar) o lo devuelvan a Norman con el motivo. |
| 1.13.3 | Puntos de miles mientras se escribe en todos los montos (Autorizado por Norman, Programar pagos, abono o nota de crédito en la factura, Nueva factura e intereses en Por aplicar); la coma es para céntimos y un punto tecleado se toma como coma. Corrección: «Registrar abono o nota de crédito» en la factura no se guardaba (el script rechazaba la operación); ahora se registra como un pago aplicado de esa factura, sin cambiar Apps Script, y el comprobante es obligatorio. |
| 1.13.2 | En «Autorizado por Norman», debajo del monto se ve en vivo si es pago total o abono y cuánto saldo queda. El campo de fecha ya no se sale del ancho en iPhone. |
| 1.13.1 | En el celular se ve la hora de sincronización («● Sincronizado 20:52») debajo del usuario. «GRUPO BSE» siempre en una línea; encabezado ajustado para teléfonos angostos. |
| 1.13.0 | Pago «Autorizado por Norman»: desde la ventana de la factura se paga o abona de inmediato lo que Norman indica de palabra, con `.txt` solo para ese pago (o transferencia manual) y comprobante. Por dentro es un pago programado y aplicado a la vez, con nota «Autorizado por Norman». El Tablero muestra los de los últimos 30 días y los movimientos llevan la etiqueta. Si la factura ya estaba programada, avisa y manda a «Por aplicar». |
| 1.12.2 | Escritura más ágil en el celular: al buscar solo se actualiza la lista y el campo de búsqueda no se vuelve a dibujar (el teclado no se reinicia); búsqueda sin autocorrección; la tecla «Buscar» muestra los resultados y cierra el teclado. |
| 1.12.1 | «Grupo BSE» en grande en el encabezado, con «Cuentas por Pagar» debajo. |
| 1.12.0 | Importación del reporte de GTI (CSV sin encabezado con punto y coma, o el Excel tal como se descarga): número de factura tomado del consecutivo, sociedad por cédula, proveedor por cédula o nombre con confirmación, notas de crédito fuera, y doble control de repetidas. La cédula queda en las observaciones para reconocer al proveedor la próxima vez. |
| 1.11.0 | Logo de Santa Eduviges en el encabezado, en la pantalla de usuarios y como ícono de la aplicación (iPhone, Android y pestaña del navegador). |
| 1.10.0 | Tablero con colores por sociedad y por tramo, con barras; tocar una sociedad abre sus facturas. Encabezado del celular ordenado (menú en cuadrícula, usuario con color). Importar facturas desde CSV o Excel en Facturas, con reconocimiento de columnas, vista previa, detección de repetidas y creación de proveedores nuevos. |
| 1.9.0 | Diseño para celular (iPhone y Android): las tablas se ven como tarjetas, sin deslizar de lado; encabezado compacto con zona segura; campos sin zoom en iPhone; ícono para instalar en la pantalla de inicio. |
| 1.8.1 | Versión de prueba del aviso de actualización (sin cambios visibles). |
| 1.8.0 | Aviso «Hay una versión nueva · Actualizar» (lee `version.json`). Norman pide revisión desde Programar pagos; Jonathan, Fernando y Daniel la ven y la marcan como revisada en Por aplicar. Menú «Canceladas» con búsqueda. Se quita la casilla «En revisión» de la ventana de la factura. |
| 1.7.0 | Sin PIN: se entra tocando el nombre (la contraseña del servicio protege cada dispositivo). Tocar el nombre arriba vuelve a la pantalla de usuarios. Fernando con rol *aplica*. Funciona con el script 1.4: en la pestaña Usuarios, la columna *contrasenaTemporal* de cada persona debe decir `acceso` la primera vez. |
| 1.6.1 | Funciona con el script de Google ya publicado (1.4), sin cambiarlo: PIN de 4 números, modo de pruebas en la aplicación (`MODO_PRUEBAS` en `js/config.js`: tras la primera vez se entra tocando el nombre) y botón «Cambiar de usuario». |
| 1.6.0 | PIN de 4 números con bloqueo tras 5 intentos; modo de pruebas sin PIN (`MODO_PRUEBAS` en Apps Script); la aplicación abre en el Tablero; pantalla de carga en lugar de "cargar Excel"; administración de usuarios desde Configuración; usuario Fernando (administrador). |
| 1.5.1 | Dirección del servicio incluida en la aplicación: en un dispositivo nuevo solo se pide la contraseña del servicio. |
| 1.5.0 | Pantalla de inicio con las tarjetas de los usuarios. La conexión con Google se ve en Configuración, protegida con la contraseña del servicio; en la entrada solo se pide la primera vez en cada dispositivo. |
| 1.4.0 | Contraseña del servicio escrita en el código de Apps Script. La dirección del servicio sale de la pantalla de entrada y se configura una vez por computadora. Número de versión visible. |
| 1.3.0 | Hoja de Google lista para subir con todos los datos (`tools/crear-hoja.mjs`). |
| 1.2.0 | Usuario y contraseña propios para cada persona; contraseñas cifradas; sesiones. |
| 1.1.0 | Google Sheets como base de datos, sincronización y trabajo sin internet. |
| 1.0.0 | Primera versión: programar pagos, aplicar y archivo para el banco. |

> Al publicar una versión, `VERSION` en `js/config.js` y `version.json` deben decir lo mismo: así las computadoras con la versión anterior ven el aviso para actualizar.

## Cambio en Apps Script para la versión 1.14

Para guardar las notas de crédito como movimiento propio, en el script de Google busque la línea

    movimiento: function (d, op, ctx, marcar) {

y agregue justo debajo:

    op = Object.assign({}, op, { tipo: op.tipoMov || op.tipo }); // tipo del movimiento: abono, pago o nc (1.14)

Guarde e implemente una versión nueva (**Implementar → Gestionar implementaciones → lápiz → Versión: Nueva versión**). La URL no cambia.

## Flujo de trabajo

1. **Norman** entra a *Programar pagos*. Marca las facturas, escribe cuánto pagar de cada una y elige la fecha. Si el monto es menor que el saldo, queda como abono.
2. **Jonathan o Daniel** entran a *Por aplicar* y descargan el `.txt` para el banco, uno por sociedad, moneda y fecha.
3. Después de subir el archivo al banco, escriben el número de comprobante y aplican el pago.

Si dos personas aplican el mismo pago, por ejemplo una de ellas sin internet, la hoja acepta solo el primero y al segundo le avisa. La pestaña **Bitacora** registra quién hizo cada cambio y cuándo.

## Instalación (una sola vez)

### 1. Crear la hoja con los datos y el script

La hoja se crea a partir del archivo **CXP Grupo BSE.xlsx**, que ya trae todas las pestañas llenas con las facturas, movimientos, proveedores, sociedades, configuración y usuarios con contraseña temporal. Se genera con `node tools/crear-hoja.mjs <excel> <salida.xlsx>`.

1. Entre a <https://drive.google.com>, elija **Nuevo → Subir archivo** y seleccione **CXP Grupo BSE.xlsx**.
2. Haga doble clic en el archivo subido. Arriba, entre a **Archivo → Guardar como hoja de cálculo de Google**. Se crea la versión de Google Sheets; trabaje siempre en esa versión, la que tiene el ícono verde.
3. En esa hoja, entre a **Extensiones → Apps Script**.
4. En el archivo `Código.gs`, borre todo y pegue el contenido de `apps-script/Codigo.gs`.
5. Haga clic en **+** junto a "Archivos", elija **Secuencia de comandos** y llámela `Logica`. Pegue ahí el contenido de `apps-script/Logica.gs`.
6. Guarde con el ícono del disquete.
7. Arriba, elija la función **configurarInicio** y haga clic en **Ejecutar**. No borra nada: solo revisa que estén todas las pestañas.
8. Google pide permisos. Haga clic en *Revisar permisos* y elija su cuenta. Si aparece "Google no verificó esta app", entre a *Configuración avanzada → Ir a … (no seguro)* y después a *Permitir*.
9. La **contraseña del servicio** se escribe en el propio código, en la línea `var CLAVE_SERVICIO = '...'` al inicio de `Codigo.gs`. Es una sola para todas las computadoras y se pide junto con el enlace `/exec` al conectar cada una. (En GitHub el código solo tiene el texto de ejemplo; la contraseña real vive únicamente en su Apps Script).
10. En la pestaña **Usuarios** está el usuario (`jonathan`, `norman`, `daniel`) y la **contraseña temporal** de cada persona. Entréguesela a cada uno en privado. La primera vez que entren, la aplicación les pide crear su propia contraseña.

> Si se empieza con una hoja vacía en lugar del archivo, `configurarInicio` crea las pestañas y los usuarios, y Jonathan carga el Excel desde la aplicación.

### 2. Publicar el servicio

1. En Apps Script, entre a **Implementar → Nueva implementación**.
2. En el engranaje de "Seleccionar tipo", elija **Aplicación web**.
3. En *Ejecutar como*, elija **Yo**. En *Quién tiene acceso*, elija **Cualquier usuario**. Esta opción solo permite llamar al servicio: para leer o escribir datos hay que iniciar sesión con usuario y contraseña.
4. Haga clic en **Implementar** y copie la **URL de la aplicación web**. Termina en `/exec`.

> Para cambiar la contraseña del servicio: edite `CLAVE_SERVICIO` en el código, guarde e implemente una nueva versión. Después hay que volver a conectar cada computadora.

> Si más adelante cambia el script, entre a **Implementar → Gestionar implementaciones → editar (lápiz) → Versión: Nueva versión**. Así se conserva la misma URL.

### 3. Primer ingreso

1. Abra la aplicación: <https://jcerros1492.github.io/cxp-grupo-bse/>.
2. La primera vez en cada computadora aparece **Configurar esta computadora**: pegue la URL del servicio y la contraseña del servicio. Quedan guardadas solo en esa computadora (no en GitHub). Sin la contraseña del servicio, el enlace no sirve para nada.
3. Entre con su usuario y la contraseña temporal. La aplicación le pide crear su contraseña.
4. Revise en *Configuración* el código de cliente del banco y el IBAN de cada sociedad. En *Proveedores*, complete el IBAN de los proveedores que falten.

## La hoja de Google

| Pestaña | Contenido |
|---|---|
| Facturas | Una fila por factura, con monto, pagado, saldo y estado actualizados |
| Movimientos | Pagos, abonos, notas de crédito e intereses de cada factura |
| Programaciones | Lo que Norman marcó para pagar y en qué quedó: programado, aplicado o devuelto |
| Proveedores / Sociedades | Catálogos con cuentas IBAN |
| Lotes | Cada pago aplicado al banco, con su comprobante |
| Bitacora | Quién hizo cada cambio y cuándo, incluidos los rechazados |
| Usuarios | Usuario, nombre, rol (`admin`, `aplica` o `aprueba`) y activo. La contraseña se guarda cifrada (columnas *hash* y *sal*); nunca se ve la real. |
| Sesiones | Computadoras con sesión abierta. Borrar una fila cierra esa sesión. |

**Administrar usuarios:** desde la aplicación, en *Configuración → Usuarios* (administradores): crear usuarios, generar un PIN temporal nuevo y activar o desactivar. También se puede desde la hoja:

- **Agregar a alguien:** escriba una fila nueva en *Usuarios* con usuario, nombre, rol, `TRUE` y una contraseña temporal en *contrasenaTemporal*.
- **Si alguien olvida su contraseña:** escriba una contraseña temporal nueva en su fila. Al entrar con ella, la aplicación le pide crear otra.
- **Quitar el acceso:** ponga `FALSE` en *activo*. Sus sesiones dejan de funcionar en la siguiente sincronización.

**No edite a mano** las pestañas de datos: los cambios deben hacerse desde la aplicación para que se validen. Leerlas, filtrarlas o copiarlas sí se puede.

## Archivo para el banco

Cada línea mide 214 caracteres y termina en CRLF. El formato reproduce byte por byte el archivo real.

| Línea | Posiciones | Contenido |
|---|---|---|
| Encabezado | 1–15 | Código de cliente con ceros a la izquierda |
| | 16–18 | Cantidad de líneas de detalle |
| | 19–26 | Fecha en formato DDMMAAAA |
| | 27–38 | Campo de control (por confirmar con el banco; se configura en la app) |
| | 51–56 | `000000` |
| Detalle | 4–25 | IBAN |
| | 26–28 | `024` para el débito, `022` para el crédito |
| | 33–52 | Monto en centavos |
| | 65–68 | Código de cliente (solo en el débito) |
| | 85+ | Descripción |

## Desarrollo

| Archivo | Contenido |
|---|---|
| `js/logica.js` | Reglas de negocio. **Se copia tal cual a `apps-script/Logica.gs`**; hay una prueba que verifica que sean iguales. |
| `apps-script/Codigo.gs` | Servicio en Google: valida con bloqueo, guarda en la hoja y devuelve los cambios. |
| `js/sync.js` | Copia local (IndexedDB), cola de cambios sin internet y sincronización. |
| `js/store.js`, `js/app.js` | Consultas y pantallas. |
| `js/importer.js` | Lectura del Excel. |
| `js/banco.js` | Archivo `.txt` del banco. |
| `sw.js` | Guarda la aplicación para que abra sin internet. |

- **Pruebas:** `node --test`.
- **Probar sin Google:** `node tests/hoja-simulada.mjs` levanta una hoja simulada en `http://localhost:8766` que ejecuta el mismo `Codigo.gs`. Los usuarios y contraseñas temporales se consultan en `/claves`.
- **No suba** Excel, `.txt` del banco ni respaldos al repositorio, porque contienen datos reales.
