# UNAM Coursera Script

Script de Google Apps Script para automatizar la importación
de datos analíticos de Coursera hacia Google Sheets.

## Descripción

La UNAM ofrece cursos MOOC a través de Coursera. Mensualmente
se descarga un archivo Excel con datos analíticos de cada curso.
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

### Fase 2 — Lectura del archivo Excel
- Acceso a carpeta en Google Drive.
- Validación de archivos disponibles.
- Conversión temporal de Excel a Google Sheets para lectura.
- Extracción de encabezados y datos de las 25 columnas.
- Registro de eventos en Log_Importación.

## Tecnologías
- Google Apps Script (JavaScript)
- Google Sheets
- Google Drive API v2

## Uso
1. Subir el archivo Excel de Coursera a la carpeta
   `UNAM_Coursera_Analiticas/Archivos_Mensuales` en Drive.
2. Actualizar las celdas B1 (mes) y B2 (año) en la pestaña
   Configuración del archivo de Sheets.
3. Ejecutar el script desde el editor de Apps Script.