// Pegar este código en el Google Sheet: Extensiones > Apps Script.
// Guarda cada envío del formulario como una fila nueva.
// Después de pegarlo: Implementar > Gestionar implementaciones > editar (lápiz)
// > Versión: "Nueva versión" > Implementar. Así la URL no cambia.

function doPost(e) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

  if (hoja.getLastRow() === 0) {
    hoja.appendRow(['Fecha', 'Nombre', 'Tipo de contacto', 'Contacto']);
  }

  hoja.appendRow([
    new Date(),
    e.parameter.nombre,
    e.parameter.tipo_contacto,
    "'" + e.parameter.contacto // el apóstrofo evita que Sheets borre ceros o el signo +
  ]);

  return ContentService.createTextOutput('ok');
}
