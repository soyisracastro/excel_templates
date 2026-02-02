# Guía de Uso: Exportador Masivo DIOT

Este módulo es el núcleo del sistema para la generación del archivo final de carga masiva compatible con el portal del SAT. Se encarga de convertir los datos de tu hoja de cálculo al formato de texto plano (`.txt`) separado por pipes (`|`).

## 📋 Funciones Principales

- **Generación de TXT**: Exporta la hoja activa a un archivo de texto con codificación **UTF-8 (con BOM)**, asegurando que el SAT reconozca acentos y la letra Ñ.
- **Conversión Automática de Países**: Traduce nombres de países (ej: "Estados Unidos") a sus códigos oficiales **ISO ALPHA-3** (ej: "USA") requeridos por la DIOT.
- **Procesamiento de Alta Velocidad**: Utiliza procesamiento en memoria (Arrays) para manejar miles de registros en segundos.

---

## 🚀 Cómo utilizar el Exportador

### 1. Requisitos de la Hoja

Para que el exportador funcione correctamente, tu hoja de Excel debe cumplir lo siguiente:

- **Encabezados**: Deben estar en la **Fila 5**.
- **Datos**: Deben comenzar en la **Fila 6**.
- **Columna de País**: El sistema busca automáticamente la columna que tenga el título `"PAÍS O JURISDICCIÓN DE RESIDENCIA FISCAL"` para aplicar la conversión a códigos ISO.

### 2. Ejecución de la Exportación

1. Sitúate en la hoja que deseas exportar (la que contiene los datos finales).
2. Presiona `ALT + F8` o ve a **Programador > Macros**.
3. Selecciona la macro: `ExportarDIOT`.
4. Haz clic en **Ejecutar**.

### 3. Archivo Generado

El sistema creará un archivo en la misma carpeta donde se encuentra tu libro de Excel con el siguiente nombre:
`DIOT_[Nombre_de_tu_Hoja]_CargaMasiva.txt`

---

## 🌎 Catálogo de Países Inteligente

El módulo incluye un catálogo de **249 países**. No necesitas preocuparte por el código ISO; puedes escribir el nombre del país y el sistema lo convertirá:

- "ALEMANIA" ➔ `DEU`
- "ESPAÑA" ➔ `ESP`
- "ESTADOS UNIDOS (LOS)" ➔ `USA`
- "OTRO" ➔ `ZZZ`

_Nota: La búsqueda no es sensible a mayúsculas o minúsculas._

---

## 🛠️ Solución de Problemas Comunes

| Problema                         | Causa Proprobable                                                            | Solución                                                                                                                           |
| :------------------------------- | :--------------------------------------------------------------------------- | :--------------------------------------------------------------------------------------------------------------------------------- |
| **"El archivo está en uso"**     | Tienes el archivo `.txt` abierto en otra aplicación (como el Bloc de Notas). | Cierra el archivo `.txt` y vuelve a ejecutar la macro.                                                                             |
| **"No hay datos para exportar"** | La macro no detectó información a partir de la Fila 6.                       | Verifica que tus datos comiencen en la Fila 6 de la hoja activa.                                                                   |
| **No convierte un país**         | El nombre del país no coincide exactamente con el catálogo oficial.          | Revisa el archivo `ModuloExportadorDIOT.bas` para ver la lista de nombres válidos o consulta la documentación del SAT relacionada. |

---

_DIOT 2026 - Módulo de Automatización v1.2_
