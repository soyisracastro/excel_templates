# Changelog - DIOT 2026 Carga Masiva

Todos los cambios notables realizados en el proyecto para la gestión de carga masiva DIOT y procesamiento de CFDI.

## [v1.2] - 2026-02-02

### ✨ Nuevas Características

- **Lector de XML CFDI (ModuloXMLCFDI.bas)**:
  - Se implementó un nuevo motor de lectura masiva de archivos XML.
  - Soporte para **CFDI 4.0 de Ingreso (Tipo I)** y **Complementos de Pago 2.0 (Tipo P)**.
  - Consolidación automática por **RFC del emisor**, sumando montos de múltiples facturas en un solo registro.
  - **Gestión de PPD**: Vinculación inteligente de pagos diferidos, extrayendo la base gravable e impuestos directamente de los nodos de pago.
  - Generación de reporte automático en una nueva hoja denominada **"CFDI_Importados"** con diseño profesional y formato de moneda.

### 🚀 Optimizaciones de Rendimiento

- **Reescritura del Exportador DIOT (Módulo3.bas)**:
  - **Velocidad masiva**: Se cambió la lectura de celdas individual a procesamiento por **Arrays en Memoria**, reduciendo drásticamente el tiempo de ejecución en hojas con miles de registros.
  - **Diccionario Estático**: La base de datos de países ahora reside de forma persistente en memoria (`Static`), eliminando el tiempo de reconstrucción del catálogo en cada consulta.
  - **Detección Dinámica**: Identificación inteligente de columnas por encabezado, eliminando la dependencia de posiciones fijas.

### 🌎 Actualización de Catálogos

- **Base de Datos de Países**:
  - Se expandió el catálogo a **249 países** con sus respectivos códigos ISO ALPHA-3.
  - Normalización de nombres (Mayúsculas/Recorte de espacios) para evitar fallos por errores de dedo en la captura.
  - Sincronización completa con el estándar del SAT para residentes en el extranjero.

### 🛠️ Correcciones y Mejoras Técnicas

- **Error 76 (Path Not Found)**:
  - Se corrigió un error crítico donde se usaba `msoFileDialogFilePicker` (3) en lugar de `msoFileDialogFolderPicker` (4), lo que causaba que el sistema intentara procesar un archivo XML como si fuera una carpeta.
  - Implementación de **Manejo de Errores para OneDrive**: El código ahora detecta y notifica cuando una carpeta está "solo en la nube", sugiriendo al usuario la opción de "Mantener siempre en este dispositivo".
  - **Normalización de Rutas**: Limpieza automática de barras finales (`\`) que causaban fallos en la detección de directorios.
- **Gestión de Archivos**: Se añadió limpieza automática de caracteres especiales (`\ / : * ? " < > |`) en los nombres de los archivos generados.
- **Manejo de Errores**: Se implementó una verificación de archivo abierto para evitar errores de ejecución cuando el archivo `.txt` de destino está siendo usado por otro programa.
- **UTF-8 con BOM**: Asegurada la codificación correcta para que el portal del SAT reconozca caracteres especiales (acentos y letra Ñ).

---

_Generado por Antigravity AI Coding Assistant._
