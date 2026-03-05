# Notas Técnicas para Desarrolladores

## Refactor ModuloXMLCFDI.bas - Versión 2.0

---

## Resumen de Cambios Arquitectónicos

### Eliminado
- `Sub CargarXMLs()` - Macro antiguo (consolidaba automáticamente)
- `Sub ActualizarDiccionario()` - Función de consolidación inline
- `Sub EscribirEnHoja()` - Función de escritura consolidada
- `NS_PAGO20` - Ya no se procesan Pagos 2.0

### Agregado
- `Sub CargarXMLProveedores()` - Nueva macro de carga (granular: 1 XML = 1 fila)
- `Sub ConcentrarDatos()` - Nueva macro de consolidación manual
- `Sub LimpiarDatos()` - Nueva macro de limpieza con confirmación
- `Private Function CargarUUIDsExistentes()` - Dedup O(1) usando Dictionary
- `Private Function ObtenerSiguienteFila()` - Búsqueda de última fila usada

### Mantenido Sin Cambios
- `Function SeleccionarCarpeta()` - Folder picker (idéntico)
- `Private Function GetAttr()` - XPath extraction (idéntico)
- `Private Function GetNodeAttr()` - Node attribute extraction (idéntico)

---

## Constantes de Configuración

```vba
' Nombres de hojas
Private Const HOJA_PROVEEDORES As String = "Datos_Proveedores"
Private Const HOJA_CONCENTRADOS As String = "Datos_Concentrados"

' Layout de filas
Private Const FILA_ENCABEZADOS As Long = 4
Private Const FILA_INICIO_DATOS As Long = 5

' Definición de columnas (A=1, B=2, ..., P=16)
Private Const COL_RFC As Long = 1        ' A
Private Const COL_NOMBRE As Long = 2     ' B
Private Const COL_UUID As Long = 3       ' C
Private Const COL_FECHA As Long = 4      ' D
Private Const COL_FOLIO As Long = 5      ' E
Private Const COL_TIPO As Long = 6       ' F
Private Const COL_METODO As Long = 7     ' G
Private Const COL_GRAV16 As Long = 8     ' H
Private Const COL_GRAV8 As Long = 9      ' I
Private Const COL_TASA0 As Long = 10     ' J
Private Const COL_EXENTO As Long = 11    ' K
Private Const COL_DESCUENTO As Long = 12 ' L
Private Const COL_IVA16 As Long = 13     ' M
Private Const COL_IVA8 As Long = 14      ' N
Private Const COL_IVARET As Long = 15    ' O
Private Const COL_TOTAL As Long = 16     ' P
Private Const ULTIMA_COL As Long = 16
```

**Beneficio:** Cambiar la estructura de columnas requiere solo actualizar estos 16 constants en un lugar.

---

## CargarXMLProveedores() - Análisis

### Entrada
- Diálogo de carpeta (FileDialog)

### Salida
- Una fila por XML valido procesado
- MsgBox con resumen (procesados, duplicados, ignorados)

### Flujo Principal

```
1. SeleccionarCarpeta() → string folderPath
2. Validar ruta (Dir + fso.FolderExists)
3. Verificar hoja Datos_Proveedores existe
4. CargarUUIDsExistentes() → Dictionary dicUUIDs [O(1) lookup]
5. ObtenerSiguienteFila() → Long nextRow
6. Para cada archivo .xml en carpeta:
   a. Parsear MSXML2.DOMDocument.6.0
   b. Filtrar tipo = "I" o "E" solamente
   c. Verificar UUID dedup (dicUUIDs.exists)
   d. Extraer 7 datos básicos (RFC, nombre, fecha, folio, método, descuento, total)
   e. Iterar traslados globales:
      - Filtrar Impuesto="002" (solo IVA, ignorar IEPS "003")
      - Categorizar por TasaOCuota y TipoFactor
      - Acumular Base → cols H-K
      - Acumular Importe → cols M-N
   f. Iterar retenciones globales:
      - Filtrar Impuesto="002"
      - Acumular Importe → col O
   g. Escribir 16 columnas en nextRow
   h. Agregar UUID a dicUUIDs
   i. Incrementar nextRow
7. Mostrar resumen
```

### Puntos Clave

**Deduplicación (líneas 168-174):**
```vba
uuid = UCase(Trim(GetAttr(...))) ' Normalizar caso y espacios
If uuid = "" Then GoTo SiguienteArchivo ' UUID obligatorio
If dicUUIDs.exists(uuid) Then
    contDuplicados = contDuplicados + 1
    GoTo SiguienteArchivo
End If
```
> Normalizar UUID a mayúsculas es crítico: PACs generan UUIDs en mixed-case.

**Signo de Egreso (línea 188):**
```vba
signo = IIf(tipoComprobante = "E", -1, 1)
' Después: todas las bases e IVA se multiplican × signo
.Cells(nextRow, COL_GRAV16).Value = baseGrav16 * signo
```
> Los Egresos con signo -1 permiten sumas automáticas en ConcentrarDatos.

**Filtrado de IEPS (línea 202):**
```vba
If impuesto = "002" Then ' Solo IVA, ignorar IEPS "003"
```
> El XML puede tener múltiples Traslado nodes (IVA + IEPS). Este filtro es esencial.

**Acumuladores de Impuestos (líneas 191-192, 206-216):**
```vba
If tipoFactor = "Exento" Then
    baseExento = baseExento + CDbl(Val(GetNodeAttr(nodo, "Base")))
ElseIf tasaOCuota = "0.160000" Then
    baseGrav16 = baseGrav16 + CDbl(Val(GetNodeAttr(nodo, "Base")))
    ivaTrasl16 = ivaTrasl16 + CDbl(Val(GetNodeAttr(nodo, "Importe")))
' ... etc
```
> Los nodos globales Traslado ya tienen valores consolidados por tasa. No iteramos conceptos.

---

## ConcentrarDatos() - Análisis

### Entrada
- Hoja "Datos_Proveedores" con datos en fila 5+

### Salida
- Nueva hoja "Datos_Concentrados"
- Una fila por RFC único
- MsgBox con conteo de proveedores

### Flujo Principal

```
1. Validar Datos_Proveedores existe y tiene datos
2. Leer rango [FILA_INICIO_DATOS : ultimaFila, COL_RFC : ULTIMA_COL] a array Variant
3. Crear Scripting.Dictionary dicRFC
4. Para cada fila del array:
   a. RFC = Trim(array(i, COL_RFC))
   b. Si RFC vacío → skip
   c. Si dicRFC.Exists(RFC):
      - Incrementar NumOps
      - Sumar 9 columnas numéricas (H-P)
      Else:
      - Crear nuevo array con valores iniciales
      - NumOps = 1
5. Crear/recrear hoja "Datos_Concentrados"
6. Escribir encabezados (12 columnas, sin UUID ni Fecha)
7. Para cada RFC en dicRFC:
   - Escribir fila con RFC, Nombre, NumOps, 9 columnas numéricas sumadas
8. Aplicar formato (moneda $, bordes, AutoFit)
9. MsgBox resumen
```

### Estructura del Array en dicRFC

```vba
' Índices del array almacenado en dicRFC(rfc)
Array(
    0 → Nombre,         ' String
    1 → NumOps,         ' Long (conteo)
    2 → Grav16,         ' Double (suma)
    3 → Grav8,          ' Double (suma)
    4 → Tasa0,          ' Double (suma)
    5 → Exento,         ' Double (suma)
    6 → Descuento,      ' Double (suma)
    7 → IVA16,          ' Double (suma)
    8 → IVA8,           ' Double (suma)
    9 → IVARet,         ' Double (suma)
    10 → Total          ' Double (suma)
)
```

### Puntos Clave

**Lectura en array (línea 320-322):**
```vba
rngDatos = wsProveedores.Range(
    wsProveedores.Cells(FILA_INICIO_DATOS, 1),
    wsProveedores.Cells(ultimaFila, ULTIMA_COL)).Value
```
> Este patrón lee el rango completo en memoria (1-based array). Es más rápido que iterar celdas.

**Suma por RFC (líneas 357-368):**
```vba
datos(1) = datos(1) + 1  ' Incrementar NumOps
datos(2) = datos(2) + CDbl(Val(rngDatos(i, COL_GRAV16)))
' ... 8 líneas más de sumas
```
> Se accede al array almacenado, se incrementa cada campo, y se reasigna.

**Egresos en la suma (implicit):**
> Como CargarXMLProveedores escribe Egresos con signo -1, las sumas aquí naturalmente restan.
> Ejemplo: Ingreso +1000 + Egreso -500 = 500 neto.

---

## LimpiarDatos() - Análisis

### Entrada
- Confirmación del usuario (vbYesNo)

### Salida
- Datos borrados (FILA_INICIO_DATOS : ultimaFila)
- Hoja "Datos_Concentrados" eliminada

### Flujo Principal

```
1. MsgBox "¿Desea limpiar?" (vbYesNo)
2. Si vbNo → Exit Sub
3. Si vbYes:
   a. Reference to Datos_Proveedores
   b. Encontrar ultimaFila
   c. Si ultimaFila >= FILA_INICIO_DATOS:
      - ClearContents [FILA_INICIO_DATOS : ultimaFila, COL_RFC : ULTIMA_COL]
   d. Si Datos_Concentrados existe:
      - Delete Datos_Concentrados sheet
   e. Activate Datos_Proveedores
   f. Select cell FILA_INICIO_DATOS
4. MsgBox "Limpieza completada"
```

### Puntos Clave

**On Error Resume Next / GoTo 0 pattern (líneas 462-464):**
```vba
On Error Resume Next
Set wsProveedores = ThisWorkbook.Sheets(HOJA_PROVEEDORES)
On Error GoTo 0
```
> Si la hoja no existe, `Set` no lanza error (thanks to Resume Next).
> Luego verificamos `If Not wsProveedores Is Nothing`.

**ClearContents vs Delete (línea 473):**
```vba
.ClearContents ' Solo borra datos, mantiene formato
```
> NOT `.Delete` (que borra también las filas).

---

## Funciones Auxiliares

### CargarUUIDsExistentes(ws) → Dictionary

**Propósito:** Cargar UUIDs existentes en O(1) para dedup rápida.

```vba
1. ultimaFila = ws.Cells(ws.Rows.Count, COL_UUID).End(xlUp).Row
2. Si ultimaFila < FILA_INICIO_DATOS:
     Retorna Dictionary vacío
3. Leer rango [FILA_INICIO_DATOS : ultimaFila, COL_UUID] a array
4. Para cada UUID en array:
   - UCase + Trim
   - Si no vacío y no existe en dic:
       dic.Add uuid, 1
5. Retorna dic
```

**Complejidad:** O(n) lectura + O(n) inserción = O(n) total
**Lookup después:** O(1) por dicUUIDs.Exists()

### ObtenerSiguienteFila(ws) → Long

**Propósito:** Encontrar la siguiente fila disponible para append.

```vba
1. ultimaFila = ws.Cells(ws.Rows.Count, COL_RFC).End(xlUp).Row
2. Si ultimaFila < FILA_INICIO_DATOS:
     Retorna FILA_INICIO_DATOS (primera carga)
3. Else:
     Retorna ultimaFila + 1 (append)
```

---

## Manejo de Errores

### Patrón Global

```vba
On Error GoTo ErrorHandler
' ... código principal ...
GoTo Cleanup

ErrorHandler:
    errNum = Err.Number
    errDesc = Err.Description
    ' Cleanup manual aquí
    MsgBox "Error: " & errNum & " - " & errDesc
    ' Fin (sin Resume)

Cleanup:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    ' ...
End Sub
```

**Razón del patrón:**
- `Resume` limpia Err.Number, así que capturamos antes
- Cleanup siempre se ejecuta (no es un `Finally`, pero funciona)
- ErrorHandler salva el estado de Err antes de cualquier otra operación

### Errores Esperados

1. **Carpeta no accesible** (líneas 78-83):
   ```vba
   On Error Resume Next
   carpetaExiste = Dir(folderPath, vbDirectory) <> "" Or fso.FolderExists(folderPath)
   On Error GoTo ErrorHandler
   ```
   > Dual check: `Dir()` es VBA nativo, `fso.FolderExists()` es API COM.

2. **XML con parseError** (línea 154):
   ```vba
   If xmlDoc.parseError.ErrorCode <> 0 Then GoTo SiguienteArchivo
   ```
   > XML inválido se salta silenciosamente (se cuenta como "ignorado").

3. **Hoja no existe** (línea 124):
   ```vba
   Set wsProveedores = ThisWorkbook.Sheets(HOJA_PROVEEDORES)
   If wsProveedores Is Nothing Then
       MsgBox "...", vbCritical, "Hoja no encontrada"
       Exit Sub
   End If
   ```

---

## XPath Queries Validadas

| Descripción | XPath | Atributo(s) |
|-------------|-------|-----------|
| Tipo comprobante | `/cfdi:Comprobante` | `TipoDeComprobante` |
| Emisor | `/cfdi:Comprobante/cfdi:Emisor` | `Rfc`, `Nombre` |
| Comprobante meta | `/cfdi:Comprobante` | `Fecha`, `Serie`, `Folio`, `MetodoPago`, `Descuento`, `Total` |
| UUID | `/cfdi:Comprobante/cfdi:Complemento/tfd:TimbreFiscalDigital` | `UUID` |
| Traslados globales | `/cfdi:Comprobante/cfdi:Impuestos/cfdi:Traslados/cfdi:Traslado` (iterate) | `Impuesto`, `TipoFactor`, `TasaOCuota`, `Base`, `Importe` |
| Retenciones globales | `/cfdi:Comprobante/cfdi:Impuestos/cfdi:Retenciones/cfdi:Retencion` (iterate) | `Impuesto`, `Importe` |

**Nota:** Todos validados contra los 76 XMLs en la carpeta `xml/` del proyecto.

---

## Performance

### Optimizaciones Aplicadas

1. **Array-based reading** (no cell-by-cell)
   ```vba
   rngDatos = ws.Range(...).Value ' Una operación
   ' vs
   For i = 1 To rows
       value = ws.Cells(i, col).Value ' rows operaciones
   ```

2. **Dictionary for dedup** (O(1) lookup vs O(n) scan)
   ```vba
   dicUUIDs.Exists(uuid) ' O(1)
   ' vs
   ws.Columns(COL_UUID).Find(uuid) ' O(n)
   ```

3. **Screen updating OFF** durante procesamiento
   ```vba
   Application.ScreenUpdating = False
   Application.Calculation = xlCalculationManual
   ```

### Benchmark Estimado

Para 100 XMLs:
- Lectura y parseo: ~50-100 ms por XML
- Escritura a hoja: ~1-2 ms por fila
- Concentración 100 filas → 15 RFCs: ~50 ms
- **Total: 5-10 segundos aprox.**

---

## Testing Checklist

- [ ] Crear hoja "Datos_Proveedores" con encabezados fila 4
- [ ] Cargar XMLs desde carpeta con XMLs válidos (76 en `xml/`)
- [ ] Verificar desglose correcto de IVA (16%, 8%, 0%, exento)
- [ ] Cargar misma carpeta dos veces → verificar UUID dedup
- [ ] Cargar carpeta A, luego carpeta B → verificar append (no reemplazo)
- [ ] Cargar XMLs con Egresos → verificar valores negativos
- [ ] Concentrar datos → verificar sumas correctas
- [ ] Limpiar datos → verificar que fila 5+ queda vacía
- [ ] Verificar que Datos_Concentrados se crea/recrea correctamente
- [ ] Probar con OneDrive (si aplica) → detectar URLs y convertir a ruta local

---

## Cambios Futuros Posibles

1. **Exportar a CSV:**
   Agregar botón que exporte "Datos_Concentrados" a pipe-delimited TXT (como ModuloExportadorDIOT).

2. **Filtrar por fecha:**
   Agregar input box en CargarXMLProveedores para cargar solo XMLs de cierto rango de fechas.

3. **Validar RFC:**
   Validar formato RFC antes de escribir (actualmente se escribe tal cual viene del XML).

4. **Importar desde CSV:**
   Leer XMLs desde un CSV con rutas locales (en lugar de carpeta única).

5. **Integración con ModuloExportadorDIOT:**
   Auto-copiar columnas de "Datos_Concentrados" a la plantilla de exportación.

---

**Documento versión:** 2.0
**Última actualización:** Febrero 2025
**Autor del refactor:** [Tu nombre aquí]
