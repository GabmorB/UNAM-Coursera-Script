function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('MOOC Coursera')
    .addItem('Subir archivo datos mensuales a sheets', 'escribirDatosMensuales')
    .addSeparator()
    .addItem('Verificar si hay archivo en carpeta de Drive', 'leerArchivoExcel')
    .addItem('Mostrar número de cursos en archivo', 'relacionarCursos')
    .addToUi();
}

const CARPETA_RAIZ = "UNAM_Coursera_Analiticas_Test";

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

function escribirDatosMensuales() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet();
  const log = hoja.getSheetByName("Log_Importación");

  // Buscar archivo en las tres carpetas
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

  // Extraer período del nombre del archivo
  const nombreArchivo = archivo.getName().toLowerCase().trim();
  const partes = nombreArchivo.split("_");

  if (partes.length < 2) {
    SpreadsheetApp.getUi().alert("Nombre de archivo incorrecto.\nFormatos válidos:\n- enero_2025\n- trimestre1_2025\n- anual_2025");
    log.appendRow([new Date(), "ERROR", "Nombre de archivo con formato incorrecto: " + nombreArchivo]);
    return;
  }

  const periodo = partes[0];
  const anio = partes[1];

  // Detectar tipo de período y pestaña destino
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

  // Verificar duplicados
  const datosExistentes = destino.getDataRange().getValues();
  for (let i = 1; i < datosExistentes.length; i++) {
    if (datosExistentes[i][0].toString().trim().toLowerCase() === periodo &&
        datosExistentes[i][1].toString().trim() === anio) {
      SpreadsheetApp.getUi().alert("Ya existen datos para " + periodo + " " + anio + " en " + pestanaDestino + ". Importación cancelada.");
      log.appendRow([new Date(), "ADVERTENCIA", "Ya existen datos para " + periodo + " " + anio + " en " + pestanaDestino]);
      return;
    }
  }

  // Leer datos del archivo
  const hojaDatos = SpreadsheetApp.openById(archivo.getId()).getSheets()[0];
  const datosExcel = hojaDatos.getDataRange().getValues();

  // Preparar filas
  const filasParaEscribir = [];
  for (let i = 1; i < datosExcel.length; i++) {
    const fila = datosExcel[i];
    if (fila[1] === "" || fila[1] === null) continue;
    filasParaEscribir.push([periodo, anio, ...fila.slice(1)]);
  }

  // Escribir de una sola vez
  const ultimaFila = destino.getLastRow() + 1;
  destino.getRange(ultimaFila, 1, filasParaEscribir.length, filasParaEscribir[0].length)
    .setValues(filasParaEscribir);

  const mensaje = "Importación completada\n" + filasParaEscribir.length + " cursos importados\nPeríodo: " + periodo + " " + anio + "\nDestino: " + pestanaDestino + "\nCarpeta: " + carpetaUsada;
  console.log(mensaje);
  log.appendRow([new Date(), "OK", "Importación completada: " + filasParaEscribir.length + " cursos | Período: " + periodo + " " + anio + " | Destino: " + pestanaDestino]);
  SpreadsheetApp.getUi().alert(mensaje);
}