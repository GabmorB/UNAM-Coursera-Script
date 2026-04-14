# UNAM Coursera Script

Script de Google Apps Script para automatizar la importación
de datos analíticos de Coursera hacia Google Sheets.

## Descripción

La UNAM ofrece cursos MOOC a través de Coursera. Mensualmente
se descarga un archivo de analíticos con datos de cada curso.
Este script automatiza la importación, almacenamiento histórico
y consulta de esos datos sin trabajo manual.

## Estructura del proyecto

El sistema opera sobre un archivo de Google Sheets con las
siguientes pestañas:

- **MOOC Coursera** — datos estructurales fijos de cada curso.
- **Datos_Mensuales** — historial acumulado mes a mes.
- **Log_Importación** — registro de errores e importaciones.
- **Configuración** — parámetros de ejecución del script.

## Fases de desarrollo

### Fase 1 — Lectura de configuración 
- Conexión al archivo de Google Sheets.
- Lectura de parámetros desde la pestaña Configuración
(mes, año y carpeta de Drive).

### Fase 2 — Lectura del archivo de analíticos 
- Acceso a carpeta en Google Drive.
- Validación de archivos disponibles.
- Apertura directa del archivo en formato Google Sheets.
- Extracción de encabezados y datos de las 25 columnas.
- Registro de eventos en Log_Importación.
- Variable global `CARPETA_RAIZ` para centralizar
la configuración de rutas.

### Fase 3 — Relación de cursos 
- Cruce de nombres de cursos entre el archivo de Coursera
y la tabla estructural de MOOC Coursera.
- Detección y registro de incompatibilidades en Log_Importación.
- Validación exitosa: 154 cursos relacionados, 0 incompatibilidades
tras corrección de nombres.

### Fase 4 — Escritura en Datos_Mensuales 
- Preparación de filas en memoria antes de escribir.
- Escritura optimizada con setValues en una sola llamada
a la API — reducción de tiempo de 2 minutos a segundos.
- Filtrado automático de filas de totales generadas
por Coursera (filas sin nombre de curso).
- Prevención de duplicados por mes y año antes de escribir.
- Registro de importación completada en Log_Importación.


## Tecnologías
- Google Apps Script (JavaScript)
- Google Sheets
- Google Drive

## Uso
1. Abrir el archivo de analíticos de Coursera en Google Sheets
   y guardarlo en formato Google Sheets.
2. Subirlo a la carpeta `Archivos_Mensuales` en Drive.
   Solo debe haber un archivo a la vez.
3. Actualizar las celdas B1 (mes) y B2 (año) en la pestaña
   `Configuración` del archivo de Sheets.
4. Ejecutar la función `escribirDatosMensuales` desde
   el editor de Apps Script.