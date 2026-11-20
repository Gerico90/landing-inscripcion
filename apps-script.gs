// Pegar este código en el Google Sheet: Extensiones > Apps Script.
// Guarda cada envío del formulario como una fila nueva.
// Después de pegarlo: Implementar > Gestionar implementaciones > editar (lápiz)
// > Versión: "Nueva versión" > Implementar. Así la URL no cambia.

const ZONA_HORARIA = 'America/Monterrey'; // hora de Saltillo
const ENCABEZADOS = ['Fecha', 'Hora', 'Nombre', 'Tipo de contacto', 'Contacto', 'Contactado'];

function doPost(e) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const ahora = new Date();

  // Mantiene la fila de encabezados al día
  hoja.getRange(1, 1, 1, ENCABEZADOS.length).setValues([ENCABEZADOS]);

  hoja.appendRow([
    Utilities.formatDate(ahora, ZONA_HORARIA, 'dd/MM/yyyy'),
    Utilities.formatDate(ahora, ZONA_HORARIA, 'HH:mm:ss'),
    e.parameter.nombre,
    e.parameter.tipo_contacto,
    "'" + e.parameter.contacto // el apóstrofo evita que Sheets borre ceros o el signo +
  ]);

  // Casilla de verificación al final del registro
  hoja.getRange(hoja.getLastRow(), ENCABEZADOS.length).insertCheckboxes();

  return ContentService.createTextOutput('ok');
}
