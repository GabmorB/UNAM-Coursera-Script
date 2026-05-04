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
- **Consulta_Mensual** — dashboard comparativo entre períodos mensuales.
- **Consulta_Curso** — evolución de un curso específico en el tiempo.
- **Ficha técnica de cursos** — datos técnicos de cada curso.
- **Catálogo de expertos** — información de instructores.
- **Lista por área** — filtro de cursos por área de conocimiento.
- **Dashboard** — métricas generales del catálogo.

## Estructura de carpetas en Google Drive

UNAM_Coursera_Analiticas/
├── Archivos_Mensuales/
├── Archivos_Trimestrales/
└── Archivos_Anuales/

El archivo debe nombrarse según el período antes de subirlo.
Drive convierte automáticamente el Excel a Google Sheets:
- Mensual: `enero_2025`, `febrero_2025`, etc.
- Trimestral: `trimestre1_2025`, `trimestre2_2025`, etc.
- Anual: `anual_2025`

## Fases de desarrollo

### Fase 1 — Lectura de configuración ✅
- Conexión al archivo de Google Sheets.
- Lectura de parámetros desde la pestaña Configuración.

### Fase 2 — Lectura del archivo de analíticos ✅
- Acceso a carpetas en Google Drive.
- Validación de archivos disponibles.
- Apertura directa del archivo en formato Google Sheets.
- Extracción de encabezados y datos de las 25 columnas.
- Registro de eventos en Log_Importación.
- Variable global `CARPETA_RAIZ` para centralizar rutas.

### Fase 3 — Relación de cursos ✅
- Cruce de nombres de cursos entre el archivo de Coursera
y la tabla estructural de MOOC Coursera.
- Detección y registro de incompatibilidades en Log_Importación.
- Validación exitosa: 154 cursos relacionados, 0 incompatibilidades
tras corrección de nombres.

### Fase 4 — Escritura de datos ✅
- Detección automática del tipo de período por nombre de archivo.
- Escritura automática en la pestaña correcta:
  Datos_Mensuales, Datos_Trimestrales o Datos_Anuales.
- Escritura optimizada con setValues en una sola llamada a la API.
- Filtrado automático de filas de totales de Coursera.
- Prevención de duplicados por período y año.
- Registro de importación completada en Log_Importación.
- Conversión automática de Excel a Google Sheets via Drive.

### Fase 5 — Dashboard comparativo ✅ (parcial)
- Pestaña Consulta_Mensual con dos selectores de período.
- Métricas comparativas: Total Enrollments, Total Completions,
  Complete Rate promedio, Avg Star Rating promedio y
  número de cursos.
- Columnas de diferencia y variación porcentual automáticas.
- Top 5 cursos por enrollments del período seleccionado.
- Pendiente: Consulta_Trimestral y Consulta_Anual.

### Fase 6 — Mejoras de usabilidad ✅
- Menú personalizado en Sheets via onOpen().
- Alertas en pantalla para confirmaciones y errores.
- Eliminación de configuración manual de mes y año.
- Tres carpetas de Drive según tipo de período.

### Fase 7 — Consulta por curso específico ✅ (parcial)
- Pestaña Consulta_Curso con desplegable de cursos.
- Muestra evolución mensual del curso seleccionado.
- Métricas: Total Enrollments, Total Completions,
  Complete Rate y Avg Star Rating por período.
- Pendiente: integrar datos trimestrales y anuales.

### Fase 8 — Dashboards trimestrales y anuales (próximamente)
- Pestaña Consulta_Trimestral — comparativa entre trimestres.
- Pestaña Consulta_Anual — comparativa entre años completos.
- Integración de datos trimestrales y anuales en Consulta_Curso.

### Fase 9 — Producción y diseño (próximamente)
- Migración de carpeta de prueba a carpeta oficial del equipo.
- Aplicar guía de estilo del Gantt MOOC 2020 (CUAIEED).
- Función principal unificada para ejecutar flujo completo.

## Tecnologías
- Google Apps Script (JavaScript)
- Google Sheets
- Google Drive

## Uso
1. Descargar el archivo de analíticos desde Coursera.
2. Renombrarlo según el formato: `periodo_año`
   (ejemplo: `enero_2025`, `trimestre2_2025`, `anual_2024`).
3. Subirlo a la carpeta correspondiente en Drive.
   Solo debe haber un archivo a la vez en total entre
   las tres carpetas.
4. Ejecutar desde el menú **MOOC Coursera** en Google Sheets
   → **Importar datos a Sheets**.

## Notas técnicas
- Las fórmulas usan referencias absolutas con `$` para evitar
  desplazamiento de rangos al agregar nuevos datos.
- La validación de duplicados usa `.toString().trim()` para
  evitar conflictos entre números y texto.
- `setValues` en lugar de `appendRow` para escritura optimizada.
- El script busca en las tres carpetas simultáneamente — solo
  puede haber un archivo en total entre las tres carpetas.