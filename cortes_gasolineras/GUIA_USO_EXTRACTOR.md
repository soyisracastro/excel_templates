# Guía de Uso: Extractor de Reportes de Gasolineras

Este documento explica cómo utilizar el script `extractor.py` para procesar los reportes mensuales de ventas.

## 1. Descripción del Script

El script **`extractor.py`** automatiza la consolidación de información de los archivos de Excel que contienen una hoja por cada día del mes.

**Características principales:**

- **Auto-detección**: Escanea la carpeta `input/` y extrae mes, sucursal y año del nombre de cada archivo. No requiere editar el script.
- **Detección Dinámica**: Localiza las tablas (VALES, TARJETAS, GASTOS) buscando sus títulos, sin importar si cambian de fila o posición (solución para caso Ajuchitlán).
- **Días por mes**: Calcula automáticamente los días del mes (28, 29, 30 o 31) según el año.
- **Limpieza de Datos**: Ignora encabezados repetidos y filas vacías.
- **Consolidación**: Genera un solo archivo por sucursal con 4 hojas detalladas.

---

## 2. Estructura de Carpetas

```
cortes_gasolineras/
├── extractor.py
├── input/                ← Colocar aquí los reportes Excel
│   ├── REPORTE-ENERO-SAN_AGUSTIN-2026.xlsx
│   ├── REPORTE-ENERO-AJUCHITLAN-2026.xlsx
│   └── ...
└── output/               ← Se genera automáticamente con los consolidados
    ├── CONSOLIDADO_SAN_AGUSTIN_ENERO_2026.xlsx
    ├── CONSOLIDADO_AJUCHITLAN_ENERO_2026.xlsx
    └── ...
```

---

## 3. Nombre de los Archivos (IMPORTANTE)

Los archivos de Excel **deben** seguir este formato:

```
REPORTE-MES-SUCURSAL-AÑO.xlsx
```

| Parte      | Descripción                              | Ejemplo         |
|------------|------------------------------------------|-----------------|
| REPORTE    | Prefijo obligatorio                      | `REPORTE`       |
| MES        | Nombre del mes en español                | `FEBRERO`       |
| SUCURSAL   | Nombre de la sucursal                    | `SAN_AGUSTIN`   |
| AÑO        | Año con 4 dígitos                        | `2026`          |

**Separadores válidos:** guión (`-`) o espacio.

**Ejemplos válidos:**

- `REPORTE-MARZO-SAN_AGUSTIN-2026.xlsx`
- `REPORTE-MARZO-AJUCHITLAN-2026.xlsx`
- `reporte marzo ajuchitlan 2026.xlsx` (mayúsculas/minúsculas no importan)

---

## 4. Instrucciones para Cambio de Mes

1. **Colocar archivos**: Copia los nuevos reportes de Excel a la carpeta `input/`.
2. **Ejecutar**:
   - Abre una terminal (PowerShell o CMD).
   - Navega a la carpeta del script y ejecuta:
     ```bash
     python extractor.py
     ```
3. **Revisar**: Los consolidados se generan en la carpeta `output/`.

> **Nota:** Puedes tener archivos de varios meses en `input/` al mismo tiempo. El script procesa todos los que encuentre.

---

## 5. Archivos Generados

El script crea archivos en `output/` con el nombre:
`CONSOLIDADO_[SUCURSAL]_[MES]_[AÑO].xlsx`

Cada archivo contiene:

- **RESUMEN**: Tabla con totales diarios de cada concepto.
- **CLIENTES**: Lista detallada de todas las Notas de Crédito.
- **VALES**: Lista detallada de todos los Vales.
- **GASTOS**: Lista detallada de todos los Gastos (Caja chica).

---

## 6. Solución de Problemas Comunes

- **"No se pudo parsear"**: El nombre del archivo no sigue el formato requerido. Verifica que incluya: REPORTE, mes, sucursal y año.
- **Error "Permission denied"**: Cierra los archivos de Excel si los tienes abiertos antes de ejecutar el script.
- **Validación**: Si notas que un total no cuadra, verifica en el archivo original si el título de la sección (ej. "VALES") está escrito correctamente. El script busca palabras clave como "VALES", "TARJETA DE CREDITO", "GASTOS".
