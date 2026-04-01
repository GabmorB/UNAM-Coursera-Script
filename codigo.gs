const CARPETA_RAIZ = "UNAM_Coursera_Analiticas_Test";

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

function leerArchivoExcel() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet();
  const config = hoja.getSheetByName("Configuración");
  const mes = config.getRange("B1").getValue();
  const anio = config.getRange("B2").getValue();
  const log = hoja.getSheetByName("Log_Importación");

  const carpetaRaiz = DriveApp.getFoldersByName(CARPETA_RAIZ).next();
  const carpetaMensual = carpetaRaiz.getFoldersByName("Archivos_Mensuales").next();
  const archivos = carpetaMensual.getFilesByType(MimeType.GOOGLE_SHEETS);

  let contador = 0;
  let archivo = null;

  while (archivos.hasNext()) {
    archivo = archivos.next();
    contador++;
  }

  if (contador === 0) {
    console.log("No se encontró ningún archivo en la carpeta.");
    log.appendRow([new Date(), "ERROR", "No se encontró archivo en Archivos_Mensuales"]);
    return;
  }

  if (contador > 1) {
    console.log("Hay más de un archivo. Deja solo uno en la carpeta.");
    log.appendRow([new Date(), "ADVERTENCIA", "Hay " + contador + " archivos en la carpeta. Solo debe haber uno."]);
    return;
  }

  console.log("Archivo encontrado: " + archivo.getName());
  console.log("Período: " + mes + " " + anio);
  log.appendRow([new Date(), "OK", "Archivo encontrado: " + archivo.getName() + " | Período: " + mes + " " + anio]);
}

function leerContenidoExcel() {
  const carpetaRaiz = DriveApp.getFoldersByName(CARPETA_RAIZ).next();
  const carpetaMensual = carpetaRaiz.getFoldersByName("Archivos_Mensuales").next();

  // Abrir directamente el Google Sheets sin copiar
  const archivos = carpetaMensual.getFilesByType(MimeType.GOOGLE_SHEETS);
  const archivo = archivos.next();
  const hojaDatos = SpreadsheetApp.openById(archivo.getId()).getSheets()[0];
  const datos = hojaDatos.getDataRange().getValues();

  console.log("Encabezados: " + datos[0]);
  console.log("Fila 1: " + datos[1]);
  console.log("Fila 2: " + datos[2]);
}

function relacionarCursos() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet();
  const log = hoja.getSheetByName("Log_Importación");
  const mooc = hoja.getSheetByName("MOOC Coursera");

  // Leer nombres de cursos de la tabla estructural
  const datosMooc = mooc.getDataRange().getValues();
  const nombresMooc = datosMooc.slice(1).map(fila => fila[1].toString().trim().toLowerCase());

  // Abrir directamente el Google Sheets sin copiar
  const carpetaRaiz = DriveApp.getFoldersByName(CARPETA_RAIZ).next();
  const carpetaMensual = carpetaRaiz.getFoldersByName("Archivos_Mensuales").next();
  const archivos = carpetaMensual.getFilesByType(MimeType.GOOGLE_SHEETS);
  const archivo = archivos.next();
  const hojaDatos = SpreadsheetApp.openById(archivo.getId()).getSheets()[0];
  const datosExcel = hojaDatos.getDataRange().getValues();

  // Comparar cursos
  let encontrados = 0;
  let noEncontrados = [];

  for (let i = 1; i < datosExcel.length; i++) {
    const nombreExcel = datosExcel[i][1].toString().trim().toLowerCase();
    if (nombresMooc.includes(nombreExcel)) {
      encontrados++;
    } else {
      noEncontrados.push(datosExcel[i][1]);
    }
  }

  // Registrar resultados
  console.log("Cursos encontrados: " + encontrados);
  console.log("Cursos no encontrados: " + noEncontrados.length);
  log.appendRow([new Date(), "INFO", "Cursos encontrados: " + encontrados]);

  noEncontrados.forEach(nombre => {
    console.log("No encontrado: " + nombre);
    log.appendRow([new Date(), "INCOMPATIBILIDAD", "Curso no encontrado en MOOC Coursera: " + nombre]);
  });
}


