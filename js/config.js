// Dirección del servicio de Google Sheets (la URL que termina en /exec).
// Puede quedar aquí aunque el repositorio sea público: sin la contraseña del
// servicio, que vive solo en Apps Script, la dirección no da acceso a nada.
export const API_URL = 'https://script.google.com/macros/s/AKfycbxBud0gvtSXjPFC9YKhSvlTxXSwIRgfENyvxSaW7ZwAYiHdcnRcXKa8PK9HYc4njTi9/exec';

// Personas de la pantalla de inicio la primera vez en un dispositivo. Después se usa
// la lista de la hoja de Google (pestaña Usuarios), que se administra desde Configuración.
export const USUARIOS = [
  { usuario: 'jonathan', nombre: 'Jonathan Cerros', rol: 'admin' },
  { usuario: 'fernando', nombre: 'Fernando', rol: 'aplica' },
  { usuario: 'norman', nombre: 'Norman Alpízar', rol: 'aprueba' },
  { usuario: 'daniel', nombre: 'Daniel Barquero', rol: 'aplica' },
];

// Versión de la aplicación: se muestra abajo en cada pantalla para confirmar qué versión está publicada.
export const VERSION = '1.18.1';

// Cada cuántos segundos se buscan cambios de los demás mientras la página está abierta.
export const SEGUNDOS_SINCRONIZAR = 20;
