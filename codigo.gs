const CARPETA_RAIZ = "UNAM_Coursera_Analiticas_Test";

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('MOOC Coursera')
    .addItem('Importar datos a Sheets', 'escribirDatosMensuales')
    .addSeparator()
    .addItem('Verificar archivo en Drive', 'leerArchivoExcel')
    .addItem('Mostrar cursos en archivo', 'relacionarCursos')
    .addToUi();
}

function leerArchivoExcel() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet();
  const log = hoja.getSheetByName("Log_Importación");

  const carpetaRaiz = DriveApp.getFoldersByName(CARPETA_RAIZ).next();
  const carpetas = ["Archivos_Mensuales", "Archivos_Trimestrales", "Archivos_Anuales"];

  let archivo = null;
  let contador = 0;
  let carpetaUsada = null;

  for (const nombreCarpeta of carpetas) {
    const carpeta = carpetaRaiz.getFoldersByName(nombreCarpeta).next();
    const archivosEnCarpeta = carpeta.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (archivosEnCarpeta.hasNext()) {
      archivo = archivosEnCarpeta.next();
      contador++;
      carpetaUsada = nombreCarpeta;
    }
  }

  if (contador === 0) {
    SpreadsheetApp.getUi().alert("No se encontró ningún archivo en ninguna carpeta.");
    log.appendRow([new Date(), "ERROR", "No se encontró archivo en ninguna carpeta"]);
    return;
  }

  if (contador > 1) {
    SpreadsheetApp.getUi().alert("Hay más de un archivo en las carpetas. Deja solo uno a la vez.");
    log.appendRow([new Date(), "ADVERTENCIA", "Hay " + contador + " archivos en las carpetas."]);
    return;
  }

  SpreadsheetApp.getUi().alert("Archivo encontrado: " + archivo.getName() + "\nCarpeta: " + carpetaUsada);
  log.appendRow([new Date(), "OK", "Archivo encontrado: " + archivo.getName() + " | Carpeta: " + carpetaUsada]);
}

function relacionarCursos() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet();
  const log = hoja.getSheetByName("Log_Importación");
  const mooc = hoja.getSheetByName("MOOC Coursera");

  const datosMooc = mooc.getDataRange().getValues();
  const nombresMooc = datosMooc.slice(1).map(fila => fila[1].toString().trim().toLowerCase());

  const carpetaRaiz = DriveApp.getFoldersByName(CARPETA_RAIZ).next();
  const carpetas = ["Archivos_Mensuales", "Archivos_Trimestrales", "Archivos_Anuales"];

  let archivo = null;
  let contador = 0;

  for (const nombreCarpeta of carpetas) {
    const carpeta = carpetaRaiz.getFoldersByName(nombreCarpeta).next();
    const archivosEnCarpeta = carpeta.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (archivosEnCarpeta.hasNext()) {
      archivo = archivosEnCarpeta.next();
      contador++;
    }
  }

  if (contador === 0) {
    SpreadsheetApp.getUi().alert("No se encontró ningún archivo en ninguna carpeta.");
    return;
  }

  if (contador > 1) {
    SpreadsheetApp.getUi().alert("Hay más de un archivo en las carpetas. Deja solo uno a la vez.");
    return;
  }

  const hojaDatos = SpreadsheetApp.openById(archivo.getId()).getSheets()[0];
  const datosExcel = hojaDatos.getDataRange().getValues();

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

  log.appendRow([new Date(), "INFO", "Cursos encontrados: " + encontrados]);
  noEncontrados.forEach(nombre => {
    log.appendRow([new Date(), "INCOMPATIBILIDAD", "Curso no encontrado en MOOC Coursera: " + nombre]);
  });

  SpreadsheetApp.getUi().alert(
    "Cursos encontrados: " + encontrados +
    "\nCursos no encontrados: " + noEncontrados.length
  );
}

function escribirDatosMensuales() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet();
  const log = hoja.getSheetByName("Log_Importación");

  const carpetaRaiz = DriveApp.getFoldersByName(CARPETA_RAIZ).next();
  const carpetas = ["Archivos_Mensuales", "Archivos_Trimestrales", "Archivos_Anuales"];

  let archivo = null;
  let contador = 0;
  let carpetaUsada = null;

  for (const nombreCarpeta of carpetas) {
    const carpeta = carpetaRaiz.getFoldersByName(nombreCarpeta).next();
    const archivosEnCarpeta = carpeta.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (archivosEnCarpeta.hasNext()) {
      archivo = archivosEnCarpeta.next();
      contador++;
      carpetaUsada = nombreCarpeta;
    }
  }

  if (contador === 0) {
    SpreadsheetApp.getUi().alert("No se encontró ningún archivo en ninguna de las carpetas.");
    log.appendRow([new Date(), "ERROR", "No se encontró archivo en ninguna carpeta"]);
    return;
  }

  if (contador > 1) {
    SpreadsheetApp.getUi().alert("Hay más de un archivo en las carpetas. Deja solo uno a la vez.");
    log.appendRow([new Date(), "ADVERTENCIA", "Hay " + contador + " archivos en las carpetas."]);
    return;
  }

  const nombreArchivo = archivo.getName().toLowerCase().trim();
  const partes = nombreArchivo.split("_");

  if (partes.length < 2) {
    SpreadsheetApp.getUi().alert("Nombre de archivo incorrecto.\nFormatos válidos:\n- enero_2025\n- trimestre1_2025\n- anual_2025");
    log.appendRow([new Date(), "ERROR", "Nombre de archivo con formato incorrecto: " + nombreArchivo]);
    return;
  }

  const periodo = partes[0];
  const anio = partes[1];

  const meses = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];
  let pestanaDestino;

  if (meses.includes(periodo)) {
    pestanaDestino = "Datos_Mensuales";
  } else if (periodo.startsWith("trimestre")) {
    pestanaDestino = "Datos_Trimestrales";
  } else if (periodo === "anual") {
    pestanaDestino = "Datos_Anuales";
  } else {
    SpreadsheetApp.getUi().alert("Período no reconocido: " + periodo + "\nFormatos válidos:\n- enero_2025\n- trimestre1_2025\n- anual_2025");
    log.appendRow([new Date(), "ERROR", "Período no reconocido: " + periodo]);
    return;
  }

  const destino = hoja.getSheetByName(pestanaDestino);

  const datosExistentes = destino.getDataRange().getValues();
  for (let i = 1; i < datosExistentes.length; i++) {
    if (datosExistentes[i][0].toString().trim().toLowerCase() === periodo &&
        datosExistentes[i][1].toString().trim() === anio) {
      SpreadsheetApp.getUi().alert("Ya existen datos para " + periodo + " " + anio + " en " + pestanaDestino + ". Importación cancelada.");
      log.appendRow([new Date(), "ADVERTENCIA", "Ya existen datos para " + periodo + " " + anio + " en " + pestanaDestino]);
      return;
    }
  }

  const hojaDatos = SpreadsheetApp.openById(archivo.getId()).getSheets()[0];
  const datosExcel = hojaDatos.getDataRange().getValues();

  const filasParaEscribir = [];
  for (let i = 1; i < datosExcel.length; i++) {
    const fila = datosExcel[i];
    if (fila[1] === "" || fila[1] === null) continue;
    filasParaEscribir.push([periodo, anio, ...fila.slice(1)]);
  }

  const ultimaFila = destino.getLastRow() + 1;
  destino.getRange(ultimaFila, 1, filasParaEscribir.length, filasParaEscribir[0].length)
    .setValues(filasParaEscribir);

  const mensaje = "Importación completada\n" + filasParaEscribir.length + " cursos importados\nPeríodo: " + periodo + " " + anio + "\nDestino: " + pestanaDestino + "\nCarpeta: " + carpetaUsada;
  log.appendRow([new Date(), "OK", "Importación completada: " + filasParaEscribir.length + " cursos | Período: " + periodo + " " + anio + " | Destino: " + pestanaDestino]);
  SpreadsheetApp.getUi().alert(mensaje);
}