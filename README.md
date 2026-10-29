# Cuentas por Pagar · Grupo BSE

Aplicación web para las cuentas por pagar de Santa Eduviges, La Hilda, Nosavar y Loma Bonita.

- La **base de datos es una hoja de Google Sheets**, que se puede abrir y consultar en cualquier momento.
- **Funciona sin internet.** Cada computadora guarda una copia y los cambios se envían a la hoja cuando vuelve la conexión.
- **Sincroniza entre las tres personas.** Cada 20 segundos, y en cada cambio, la aplicación trae lo que hicieron los demás.

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
9. En la pestaña **Usuarios** está el usuario (`jonathan`, `norman`, `daniel`) y la **contraseña temporal** de cada persona. Entréguesela a cada uno en privado. La primera vez que entren, la aplicación les pide crear su propia contraseña.

> Si se empieza con una hoja vacía en lugar del archivo, `configurarInicio` crea las pestañas y los usuarios, y Jonathan carga el Excel desde la aplicación.

### 2. Publicar el servicio

1. En Apps Script, entre a **Implementar → Nueva implementación**.
2. En el engranaje de "Seleccionar tipo", elija **Aplicación web**.
3. En *Ejecutar como*, elija **Yo**. En *Quién tiene acceso*, elija **Cualquier usuario**. Esta opción solo permite llamar al servicio: para leer o escribir datos hay que iniciar sesión con usuario y contraseña.
4. Haga clic en **Implementar** y copie la **URL de la aplicación web**. Termina en `/exec`.

> Si más adelante cambia el script, entre a **Implementar → Gestionar implementaciones → editar (lápiz) → Versión: Nueva versión**. Así se conserva la misma URL.

### 3. Primer ingreso

1. Abra la aplicación: <https://jcerros1492.github.io/cxp-grupo-bse/>.
2. Pegue la URL del servicio y entre con su usuario y la contraseña temporal. La aplicación le pide crear su contraseña. La URL se recuerda en esa computadora; si se escribe en `js/config.js`, ya no se pide.
3. Revise en *Configuración* el código de cliente del banco y el IBAN de cada sociedad. En *Proveedores*, complete el IBAN de los proveedores que falten.

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

**Administrar usuarios desde la hoja:**

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
