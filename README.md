# UNAM Coursera Script

Script de Google Apps Script para automatizar la importación
de datos analíticos de Coursera hacia Google Sheets.

## Descripción

La UNAM ofrece cursos MOOC a través de Coursera. Se descargan
archivos de analíticos con datos de cada curso en distintos
períodos — mensual, trimestral o anual. Este script automatiza
la importación, almacenamiento histórico y consulta de esos
datos sin trabajo manual.

## Estructura del proyecto

El sistema opera sobre un archivo de Google Sheets con las
siguientes pestañas:

- **MOOC Coursera** — datos estructurales fijos de cada curso.
- **Datos_Mensuales** — historial acumulado mes a mes.
- **Datos_Trimestrales** — historial acumulado por trimestre.
- **Datos_Anuales** — historial acumulado por año.
- **Log_Importación** — registro de errores e importaciones.
- **Consulta_Mensual** — dashboard comparativo entre períodos.

## Estructura de carpetas en Google Drive

UNAM_Coursera_Analiticas/
├── Archivos_Mensuales/
├── Archivos_Trimestrales/
└── Archivos_Anuales/

El archivo debe nombrarse según el período antes de subirlo:
- Mensual: `enero_2025`, `febrero_2025`, etc.
- Trimestral: `trimestre1_2025`, `trimestre2_2025`, etc.
- Anual: `anual_2025`

## Fases de desarrollo

### Fase 1 — Lectura de configuración 
- Conexión al archivo de Google Sheets.
- Lectura de parámetros desde la pestaña Configuración.

### Fase 2 — Lectura del archivo de analíticos 
- Acceso a carpeta en Google Drive.
- Validación de archivos disponibles.
- Apertura directa del archivo en formato Google Sheets.
- Extracción de encabezados y datos de las 25 columnas.
- Registro de eventos en Log_Importación.
- Variable global `CARPETA_RAIZ` para centralizar rutas.

### Fase 3 — Relación de cursos 
- Cruce de nombres de cursos entre el archivo de Coursera
y la tabla estructural de MOOC Coursera.
- Detección y registro de incompatibilidades en Log_Importación.
- Validación exitosa: 154 cursos relacionados, 0 incompatibilidades
tras corrección de nombres.

### Fase 4 — Escritura de datos 
- Detección automática del tipo de período por nombre de archivo.
- Escritura automática en la pestaña correcta según el tipo:
  Datos_Mensuales, Datos_Trimestrales o Datos_Anuales.
- Escritura optimizada con setValues en una sola llamada a la API.
- Filtrado automático de filas de totales de Coursera.
- Prevención de duplicados por período y año.
- Registro de importación completada en Log_Importación.
- Conversión automática de Excel a Google Sheets via Drive.

### Fase 5 — Dashboard comparativo (parcial)
- Pestaña Consulta_Mensual con dos selectores de período.
- Métricas comparativas con diferencia y variación porcentual.
- Top 5 cursos por enrollments del período seleccionado.
- Pendiente: Consulta_Trimestral y Consulta_Anual.

### Fase 6 — Mejoras de usabilidad 
- Menú personalizado en Sheets via onOpen().
- Alertas en pantalla para confirmaciones y errores.
- Eliminación de configuración manual de mes y año.

### Fase 7 — Dashboards trimestrales y anuales (próximamente)
- Pestaña Consulta_Trimestral — comparativa entre trimestres.
- Pestaña Consulta_Anual — comparativa entre años completos.
- Consulta por curso específico — evolución a lo largo del tiempo.

## Tecnologías
- Google Apps Script (JavaScript)
- Google Sheets
- Google Drive

## Uso
1. Descargar el archivo de analíticos desde Coursera.
2. Renombrarlo según el formato: `periodo_año`
   (ejemplo: `enero_2025`, `trimestre2_2025`, `anual_2024`).
3. Subirlo a la carpeta correspondiente en Drive.
   Solo debe haber un archivo a la vez en total.
4. Ejecutar desde el menú **MOOC Coursera** en Google Sheets.