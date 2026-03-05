# Changelog - DIOT 2026 Carga Masiva

Todos los cambios notables realizados en el proyecto para la gestión de carga masiva DIOT y procesamiento de CFDI.

## [v2.0] - 2025-02-07

### ✨ Nuevas Características

#### Carga Granular de XMLs
- Cambio de modelo: **Una fila por comprobante** (antes consolidado por RFC)
- Nueva macro `CargarXMLProveedores()` reemplaza la antigua `CargarXMLs()`
- Permite revisión detallada de datos antes de consolidar
- Modo "append" para cargas múltiples desde diferentes carpetas

#### Soporte Completo para Egresos
- Procesa comprobantes tipo **Egreso (E)** (notas de crédito)
- Registra con **valores negativos** para representar devoluciones/descuentos
- Se netan automáticamente en consolidación
- Útil para cálculo de IVA acreditable bajo autodeterminación

#### Desglose Detallado de IVA por Tasa
- Nuevas columnas para bases: **16%, 8%, 0%, Exento**
- Nuevas columnas para IVA: **16%, 8%** (separados)
- Extracción precisa de nodos globales `cfdi:Traslado` del XML
- IVA Retenido registrado por separado

#### Información de Referencia Ampliada
- **Fecha** (YYYY-MM-DD): Para auditoría temporal
- **Serie-Folio**: Para cruce con contabilidad
- **Tipo** (I/E): Identificación de Ingreso vs Egreso
- **Método de Pago**: PUE, PPD, etc.

#### Deduplicación Automática por UUID
- Sistema O(1) usando `Scripting.Dictionary`
- Previene cargar el mismo XML dos veces
- Soporta cargas desde múltiples carpetas sin duplicados
- Reporta contador de duplicados omitidos

#### Consolidación Manual y Flexible
- Nueva macro `ConcentrarDatos()` genera hoja separada
- Permite revisar datos detallados antes de consolidar
- Agrupa por RFC y suma automáticamente
- Formato profesional (moneda, bordes, encabezados)
- Una fila consolidada por proveedor (RFC)

#### Limpieza Segura con Confirmación
- Nueva macro `LimpiarDatos()`
- Confirmación (vbYesNo) antes de borrar
- Limpia datos y elimina hojas generadas
- Previene borradores accidentales

### 🏗️ Cambios Arquitectónicos

- **Nuevas funciones privadas:**
  - `CarcargarUUIDsExistentes()` - Dedup de O(1)
  - `ObtenerSiguienteFila()` - Búsqueda de fila append
- **Constantes de configuración:** 17 constantes para layout
- **Eliminadas:** `ActualizarDiccionario()`, `EscribirEnHoja()`, soporte Pagos (P)
- **Resultado:** ~575 líneas, mejor separación de responsabilidades

### 📊 Cambios de Modelo de Datos

| Aspecto | v1.2 | v2.0 |
|--------|------|------|
| **Granularidad** | 1 RFC = 1 fila | 1 Comprobante = 1 fila |
| **Consolidación** | Automática | Manual (botón) |
| **IVA Detalle** | No | Sí, por 4 tasas |
| **Egresos (E)** | ❌ No | ✅ Sí (negativos) |
| **Carga Múltiple** | Reemplaza | Append |
| **Dedup** | No | Sí, por UUID |
| **Hoja Resultado** | CFDI_Importados | Datos_Concentrados |

### 🐛 Correcciones

- UUID normalizadas a mayúsculas (PACs generan mixed-case)
- IEPS filtrado correctamente (Impuesto="002" solo IVA)
- Mejor manejo de campos faltantes
- Detección mejorada de OneDrive/SharePoint

### 📝 Documentación Nuevas

- `DOCUMENTACION_REFACTOR_MODULO_XML.md` - Guía completa (2000+ líneas)
- `NOTAS_ACTUALIZACION_v2.0.md` - Resumen ejecutivo para usuarios
- `GUIA_INSTALACION_BOTONES.md` - Paso a paso para instalación
- `NOTAS_TECNICAS_DESARROLLADOR.md` - Análisis arquitectónico

### ⚠️ Cambios Incompatibles

- Nueva estructura de hojas: "Datos_Proveedores" reemplaza "CFDI_Importados"
- Nuevos encabezados: 16 columnas (antes 9)
- Datos granulares no son directamente compatibles con reportes v1.2
- Requiere recrear botones (3 macros nuevas)

### 🔄 Ruta de Migración desde v1.2

1. Backup del libro anterior
2. Reemplazar ModuloXMLCFDI.bas (v2.0)
3. Crear hoja "Datos_Proveedores" con encabezados v2.0
4. Crear 3 botones (Cargar XML, Concentrar Datos, Limpiar Datos)
5. Cargar XMLs nuevamente (datos granulares)
6. Usar botón "Concentrar Datos" para resumen

---

## [v1.2] - 2025-02-02

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

---

## Documentación Actualizada en v2.0

Se agregaron los siguientes archivos de documentación para facilitar la distribución y comunicación a usuarios:

### Documentación para Usuarios
- `NOTIFICACION_USUARIOS_v2.0.txt` - Comunicado de lanzamiento (5 min de lectura)
- `NOTAS_ACTUALIZACION_v2.0.md` - Resumen ejecutivo de mejoras
- `GUIA_INSTALACION_BOTONES.md` - Instalación paso a paso de botones (20 min)
- `DOCUMENTACION_REFACTOR_MODULO_XML.md` - Guía completa y referencia (45 min)
- `DOCUMENTACION_INDICE.txt` - Índice de lectura por perfil de usuario
- `EMAIL_COMUNICADO_USUARIOS.txt` - Plantilla para comunicado por correo

### Documentación para Desarrolladores
- `NOTAS_TECNICAS_DESARROLLADOR.md` - Análisis arquitectónico profundo (60 min)
- `QA_TESTING_CHECKLIST.txt` - 25+ test cases detallados para QA

### Total de Documentación
- **~3,000 líneas** de documentación clara, accesible y bien organizada
- **~60 KB** de archivos (tamaño manejable)
- Flujos de lectura recomendados por perfil (usuario final, técnico, desarrollador)

---

_Generado por Claude Code - Anthropic._
