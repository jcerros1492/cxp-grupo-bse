# Cuentas por Pagar

Aplicación web para las cuentas por pagar de las cuatro sociedades (Santa Eduviges, La Hilda, Nosavar y Loma Bonita).

## Flujo de trabajo

1. **Norman** entra a *Programar pagos*, filtra por sociedad, moneda o proveedor, marca las facturas, escribe cuánto pagar de cada una y elige la fecha de pago. Si el monto es menor que el saldo, el pago queda registrado como abono y la factura sigue disponible para programarla de nuevo.
2. **Jonathan o Daniel** entran a *Por aplicar*. Lo programado aparece agrupado por sociedad, moneda y fecha, y cada grupo genera su archivo `.txt` para el banco.
3. Después de subir el archivo al banco, se escribe el número de comprobante y se aplica el pago. Las facturas bajan su saldo y la aplicación registra quién aplicó cada pago y cuándo.

## Cómo usarla

Abra la página publicada en GitHub Pages, elija su usuario y cargue el Excel de cuentas por pagar. El Excel **se lee dentro del navegador**: no se sube a ningún servidor ni a GitHub.

Antes de generar el primer archivo para el banco, complete la *Configuración*:

- Código de cliente del banco.
- Cuenta IBAN de cada sociedad que se debita, en colones y en dólares.
- IBAN de cada proveedor, en *Proveedores*.

## Limitación de esta primera versión

Los datos se guardan en el navegador de cada computadora (`localStorage`). Esto significa que lo que Norman programa en su computadora **no** lo ven Jonathan ni Daniel en las suyas. El siguiente paso es conectar una base de datos compartida. Todo el acceso a los datos está concentrado en `js/store.js`, así que el cambio no afecta las pantallas.

Mientras tanto, conviene descargar un respaldo con frecuencia desde *Configuración → Descargar respaldo*.

## Formato del archivo para el banco

El formato se dedujo de un archivo real y la aplicación lo reproduce byte por byte. Cada línea mide 214 caracteres y termina en CRLF.

| Línea | Posiciones | Contenido |
|---|---|---|
| Encabezado | 1–15 | Código de cliente, con ceros a la izquierda |
| | 16–18 | Cantidad de líneas de detalle (débito + créditos) |
| | 19–26 | Fecha en formato DDMMAAAA |
| | 27–38 | Campo de control (**por confirmar con el banco**; se configura) |
| | 51–56 | `000000` |
| Detalle | 1–3 | `000` |
| | 4–25 | IBAN |
| | 26–28 | `024` para el débito, `022` para el crédito |
| | 33–52 | Monto en centavos, con 20 dígitos |
| | 65–68 | Código de cliente (solo en la línea de débito) |
| | 85–214 | Descripción, sin tildes |

## Publicar en GitHub Pages

La aplicación usa solo archivos estáticos, sin compilación:

1. Vaya a *Settings → Pages*.
2. En *Source*, elija *Deploy from a branch*.
3. Seleccione la rama `main` y la carpeta `/ (root)`.

## Desarrollo

- Probar en la computadora: `python3 -m http.server 8000` y abrir <http://localhost:8000>.
- Ejecutar las pruebas: `node --test`. Las pruebas usan cuentas ficticias.
- **No suba** archivos de Excel, archivos `.txt` del banco ni respaldos `.json` al repositorio, porque contienen datos reales. El archivo `.gitignore` ya los excluye.

| Archivo | Contenido |
|---|---|
| `js/importer.js` | Lectura del Excel. Las filas negativas se convierten en abonos de su factura y el vencimiento se estima a 30 días cuando falta. |
| `js/store.js` | Datos y reglas: saldos, estados, programaciones y lotes. |
| `js/banco.js` | Generador del `.txt` para el banco. |
| `js/app.js` | Pantallas. |
| `js/vendor/xlsx.full.min.js` | SheetJS 0.18.5 (licencia Apache 2.0) para leer y escribir Excel. |
