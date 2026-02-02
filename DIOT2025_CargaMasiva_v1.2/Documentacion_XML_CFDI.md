# Guía de Uso: Lector de XML CFDI para DIOT

Este módulo permite automatizar la extracción de información desde archivos XML de Comprobantes Fiscales Digitales (CFDI) versión 4.0 para facilitar el llenado de la carga masiva DIOT.

## 📋 Características Principales

- **Procesamiento Masivo**: Lee todos los archivos XML de una carpeta seleccionada.
- **Tipos de Comprobante**: Soporta Tipo **I** (Ingreso/Facturas) y Tipo **P** (Complementos de Pago 2.0).
- **Consolidación por RFC**: Suma automáticamente montos e impuestos de múltiples facturas de un mismo emisor.
- **Vínculos PPD**: Detecta pagos de facturas con método PPD y extrae la base de IVA efectivamente pagada.

---

## 🚀 Instrucciones de Uso

### 1. Preparación

Asegúrate de tener tus archivos XML (Ingresos y Pagos) en una carpeta local de tu computadora.

### 2. Ejecución de la Macro

1. Abre el archivo de Excel `DIOT2026_CargaMasiva_v1.1.xlsm`.
2. Presiona `ALT + F8` o ve a la pestaña **Programador > Macros**.
3. Selecciona la macro llamada: `CargarXMLs`.
4. Haz clic en **Ejecutar**.

### 3. Selección de Carpetas

Se abrirá una ventana emergente. Busca y selecciona la carpeta donde guardaste tus archivos XML. Haz clic en **Aceptar**.

### 4. Revisión de Resultados

Al terminar el proceso (aparecerá un mensaje de "Proceso completado"), se creará una nueva hoja llamada **"CFDI_Importados"**.

---

## 📊 Descripción de las Columnas Generadas

| Columna                | Descripción                                                                                 |
| :--------------------- | :------------------------------------------------------------------------------------------ |
| **RFC**                | Registro Federal de Contribuyentes del Emisor.                                              |
| **Nombre**             | Razón social o nombre del proveedor.                                                        |
| **Subtotal Acum.**     | Suma de las bases gravables de todas las facturas procesadas.                               |
| **IVA Trasladado**     | Total de IVA que el proveedor te trasladó (Efectivamente pagado en caso de complementos P). |
| **IVA Retenido**       | Total de IVA retenido al proveedor (si aplica).                                             |
| **Total Acum.**        | Importe total de la operación (incluyendo impuestos).                                       |
| **Num. Facturas**      | Conteo de cuántos archivos XML se encontraron para ese RFC.                                 |
| **UUIDs Relacionados** | Lista de folios fiscales procesados para control y auditoría.                               |
| **Método Pago**        | Indica si la operación fue PUE (una exhibición) o PPD (pago diferido/parcialidades).        |

---

## ⚠️ Notas Técnicas y Recomendaciones

- **Consolidación**: Si un proveedor tiene 10 facturas en la misma carpeta, verás una sola fila con la suma de las 10, lo cual es ideal para la captura en el portal del SAT.
- **Complementos de Pago**: El sistema busca los nodos de impuestos dentro del complemento de pago. Si un pago no especifica impuestos a nivel de documento relacionado, intentará obtenerlos del nodo global de totales del pago.
- **Permisos**: Asegúrate de que los archivos XML no estén abiertos por otro programa durante el proceso.

---

_DIOT 2026 - Módulo de Automatización v1.2_
