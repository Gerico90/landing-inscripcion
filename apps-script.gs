// Pegar este código en el Google Sheet: Extensiones > Apps Script.
// Guarda cada envío del formulario como una fila nueva.
// Para publicar cambios: guardar (Ctrl + S) > Implementar > Nueva implementación > Aplicación web
// (Ejecutar como: Yo, Acceso: Cualquier usuario) > copiar la URL nueva y ponerla en SCRIPT_URL de index.html.
// Archivar la implementación anterior SOLO después de publicar la página con la URL nueva.

const ZONA_HORARIA = 'America/Monterrey'; // hora de Saltillo
const ENCABEZADOS = ['Fecha', 'Hora', 'Nombre', 'Tipo de contacto', 'Contacto', 'Contactado'];
const TIPOS_VALIDOS = ['WhatsApp', 'Correo electrónico'];

function doPost(e) {
  const datos = e.parameter;

  // Antispam: el campo trampa solo lo llenan bots. Respondemos "ok" para no darles pistas.
  if (datos.sitio_web) {
    return responder('ok');
  }

  // Validación en el servidor (un bot puede enviar directo a esta URL sin pasar por la página)
  const nombre = String(datos.nombre || '').trim();
  const tipo = String(datos.tipo_contacto || '');
  const contacto = String(datos.contacto || '').trim();

  const nombreValido = nombre.length <= 60 && /^[\p{L}][\p{L}\s'-]*[\p{L}]$/u.test(nombre);
  const tipoValido = TIPOS_VALIDOS.indexOf(tipo) !== -1;
  const contactoValido = tipo === 'WhatsApp'
    ? /^\d{10}$/.test(contacto)
    : contacto.length <= 100 && /^[^\s@]+@[^\s@]+\.[a-zA-Z]{2,}$/.test(contacto);

  if (!nombreValido || !tipoValido || !contactoValido) {
    return responder('invalido');
  }

  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const ahora = new Date();

  // Mantiene la fila de encabezados al día
  hoja.getRange(1, 1, 1, ENCABEZADOS.length).setValues([ENCABEZADOS]);

  hoja.appendRow([
    Utilities.formatDate(ahora, ZONA_HORARIA, 'dd/MM/yyyy'),
    Utilities.formatDate(ahora, ZONA_HORARIA, 'HH:mm:ss'),
    nombre,
    tipo,
    "'" + contacto // el apóstrofo evita que Sheets borre ceros o el signo +
  ]);

  // Casilla de verificación al final del registro
  hoja.getRange(hoja.getLastRow(), ENCABEZADOS.length).insertCheckboxes();

  return responder('ok');
}

// La página lee esta respuesta para saber si el registro se guardó.
function responder(texto) {
  return ContentService.createTextOutput(texto);
}
