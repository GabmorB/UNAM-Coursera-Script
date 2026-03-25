function leerConfiguracion() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet();
  const config = hoja.getSheetByName("Configuración");
  
  const mes = config.getRange("B1").getValue();
  const anio = config.getRange("B2").getValue();
  const carpeta = config.getRange("B3").getValue();
  
  console.log("Mes: " + mes);
  console.log("Año: " + anio);
  console.log("Carpeta: " + carpeta);
}