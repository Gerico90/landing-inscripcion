// Pegar este código en el Google Sheet: Extensiones > Apps Script.
// Guarda cada envío del formulario como una fila nueva.

function doPost(e) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

  if (hoja.getLastRow() === 0) {
    hoja.appendRow(['Fecha', 'Nombre', 'Teléfono', 'Motivo']);
  }

  hoja.appendRow([
    new Date(),
    e.parameter.nombre,
    "'" + e.parameter.telefono, // el apóstrofo evita que Sheets borre ceros o el signo +
    e.parameter.motivo
  ]);

  return ContentService.createTextOutput('ok');
}
