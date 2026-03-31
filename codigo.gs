function leerConfiguracion() {
  // Obtener el archivo de Sheets activo
  const hoja = SpreadsheetApp.getActiveSpreadsheet();
  
  // Acceder a la pestaña Configuración
  const config = hoja.getSheetByName("Configuración");
  
  // Leer los valores
  const mes = config.getRange("B1").getValue();
  const anio = config.getRange("B2").getValue();
  const carpeta = config.getRange("B3").getValue();
  
  // Imprimir en consola
  console.log("Mes: " + mes);
  console.log("Año: " + anio);
  console.log("Carpeta: " + carpeta);
}

function leerArchivoExcel() {
  // Leer configuración
  const hoja = SpreadsheetApp.getActiveSpreadsheet();
  const config = hoja.getSheetByName("Configuración");
  const mes = config.getRange("B1").getValue();
  const anio = config.getRange("B2").getValue();
  const log = hoja.getSheetByName("Log_Importación");

  // Acceder a la carpeta
  const carpetaRaiz = DriveApp.getFoldersByName("UNAM_Coursera_Analiticas").next();
  const carpetaMensual = carpetaRaiz.getFoldersByName("Archivos_Mensuales").next();

  // Buscar archivos Excel en la carpeta
  const archivos = carpetaMensual.getFilesByType(MimeType.MICROSOFT_EXCEL);
  
  // Contar cuántos archivos hay
  let contador = 0;
  let archivo = null;

  while (archivos.hasNext()) {
    archivo = archivos.next();
    contador++;
  }

  // Validar cantidad de archivos
  if (contador === 0) {
    console.log("No se encontró ningún archivo Excel en la carpeta.");
    log.appendRow([new Date(), "ERROR", "No se encontró archivo Excel en Archivos_Mensuales"]);
    return;
  }

  if (contador > 1) {
    console.log("Hay más de un archivo Excel. Deja solo uno en la carpeta.");
    log.appendRow([new Date(), "ADVERTENCIA", "Hay " + contador + " archivos Excel en la carpeta. Solo debe haber uno."]);
    return;
  }

  // Si hay exactamente uno, continuar
  console.log("Archivo encontrado: " + archivo.getName());
  console.log("Período: " + mes + " " + anio);
  log.appendRow([new Date(), "OK", "Archivo encontrado: " + archivo.getName() + " | Período: " + mes + " " + anio]);
}

function leerContenidoExcel() {
  // Acceder a la carpeta
  const carpetaRaiz = DriveApp.getFoldersByName("UNAM_Coursera_Analiticas").next();
  const carpetaMensual = carpetaRaiz.getFoldersByName("Archivos_Mensuales").next();

  // Obtener el archivo Excel
  const archivos = carpetaMensual.getFilesByType(MimeType.MICROSOFT_EXCEL);
  const archivo = archivos.next();

  // Convertir temporalmente a Google Sheets para leerlo
  const blob = archivo.getBlob();
  const temporal = SpreadsheetApp.openById(
    Drive.Files.insert(
      { title: "temporal_lectura", mimeType: MimeType.GOOGLE_SHEETS },
      blob
    ).id
  );

  // Leer la primera hoja del archivo
  const hojaDatos = temporal.getSheets()[0];
  const datos = hojaDatos.getDataRange().getValues();

  // Imprimir las primeras 3 filas para verificar
  console.log("Encabezados: " + datos[0]);
  console.log("Fila 1: " + datos[1]);
  console.log("Fila 2: " + datos[2]);

  // Eliminar el archivo temporal
  DriveApp.getFileById(temporal.getId()).setTrashed(true);
}