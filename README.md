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

- **MOOC Coursera** — datos estructurales fijos de cada curso (169 cursos).
- **Datos_Mensuales** — historial acumulado mes a mes.
- **Datos_Trimestrales** — historial acumulado por trimestre.
- **Datos_Anuales** — historial acumulado por año.
- **Log_Importación** — registro de errores e importaciones.
- **Consulta_Periodo** — dashboard comparativo entre dos períodos
  cualquiera (mensual, trimestral o anual) con métricas agregadas
  de todos los cursos.
- **Consulta_Curso** — comparativa de un curso específico entre
  dos períodos cualquiera (mensual, trimestral o anual).
- **Ficha técnica de cursos** — datos técnicos de cada curso.
- **Catálogo de expertos** — información de instructores.
- **Lista por área** — filtro de cursos por área de conocimiento.
- **Dashboard** — métricas generales del catálogo.

## Estructura de carpetas en Google Drive

UNAM_Coursera_Analiticas_Test/
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
- Limpieza automática de espacios en valores de texto con `.trim()`
  al momento de importar para garantizar coincidencias exactas
  en fórmulas de consulta.

### Fase 5 — Dashboard comparativo ✅
- Pestaña Consulta_Periodo con selector de tipo (Mensual/Trimestral/Anual),
  período y año para dos períodos a comparar.
- Un único selector de Tipo aplica a ambos períodos simultáneamente.
- Métricas comparativas agregadas de todos los cursos:
  Total Enrollments (suma), Total Completions (suma),
  Complete Rate promedio, AVG Star Rating promedio y
  número de cursos.
- Columna de diferencia automática entre períodos.
- Fórmulas con FILTER + SUM/AVERAGE/COUNTA consultando
  dinámicamente las tres hojas de datos según el tipo elegido.

### Fase 6 — Mejoras de usabilidad ✅
- Menú personalizado en Sheets via onOpen().
- Alertas en pantalla para confirmaciones y errores.
- Eliminación de configuración manual de mes y año.
- Tres carpetas de Drive según tipo de período.

### Fase 7 — Consulta por curso específico ✅
- Pestaña Consulta_Curso con desplegable de 169 cursos.
- Selector de Tipo (Mensual/Trimestral/Anual), Período y Año
  independientes para Período A y Período B.
- Métricas comparativas del curso seleccionado:
  Total Enrollments, Total Completions, Complete Rate
  y AVG Star Rating.
- Columna de diferencia automática entre períodos.
- Fórmulas con FILTER consultando dinámicamente las tres
  hojas de datos según el tipo elegido en cada período.
- Validado con datos mensuales, trimestrales y anuales.

### Fase 8 — Producción y diseño (próximamente)
- Migración de carpeta de prueba a carpeta oficial del equipo
  (cambiar `CARPETA_RAIZ` de `UNAM_Coursera_Analiticas_Test`
  a la carpeta oficial).
- Aplicar guía de estilo del Gantt MOOC 2020 (CUAIEED):
  colores azul marino y rojo, logos de CUAIEED y MOOC UNAM.
- Función principal unificada para ejecutar flujo completo
  con un solo botón.

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
- Los años en los desplegables de consulta se gestionan
  manualmente — agregar el año a la validación de datos
  cuando se importe un nuevo año.
- Los rangos de fórmulas están definidos hasta la fila 10000
  para soportar crecimiento de la base de datos.
- Las fórmulas de Consulta_Curso y Consulta_Periodo usan
  `VALUE()` sobre el selector de año para garantizar
  coincidencia con valores numéricos en las hojas de datos.