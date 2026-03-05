Attribute VB_Name = "ModuloInstrucciones"
Option Explicit

Private Const HOJA_INSTRUCCIONES As String = "Instrucciones"

Sub GenerarHojaInstrucciones()
    Dim ws As Worksheet
    Dim r As Long

    ' Eliminar si ya existe
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(HOJA_INSTRUCCIONES).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Crear hoja al inicio del libro
    Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    ws.Name = HOJA_INSTRUCCIONES

    ' Configurar hoja
    ws.Cells.Font.Name = "Aptos"
    ws.Cells.Font.Size = 11
    ws.Columns("A").ColumnWidth = 4
    ws.Columns("B").ColumnWidth = 85
    ActiveWindow.DisplayGridlines = False

    r = 2

    ' =============================================
    ' TITULO PRINCIPAL
    ' =============================================
    With ws.Cells(r, 2)
        .Value = "DIOT - Carga Masiva"
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = RGB(44, 62, 80)
    End With
    r = r + 1

    With ws.Cells(r, 2)
        .Value = "Declaracion Informativa de Operaciones con Terceros"
        .Font.Size = 13
        .Font.Color = RGB(127, 140, 141)
    End With
    r = r + 2

    ' Linea separadora
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 2))
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(52, 152, 219)
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    r = r + 2

    ' =============================================
    ' DESCRIPCION
    ' =============================================
    r = WriteHeader(ws, r, "Descripcion")
    r = WriteText(ws, r, "Esta plantilla automatiza la generacion del archivo TXT para la carga masiva de la DIOT ante el SAT.")
    r = WriteText(ws, r, "Cuenta con macros para: leer XMLs de CFDI 4.0, concentrar datos por proveedor y exportar en formato pipe (|).")
    r = r + 1

    ' =============================================
    ' HOJAS DEL LIBRO
    ' =============================================
    r = WriteHeader(ws, r, "Hojas del libro")
    r = WriteText(ws, r, "Este libro contiene las siguientes hojas de trabajo:")
    r = r + 1

    r = WriteBullet(ws, r, "Instrucciones", "Esta hoja. Guia de uso paso a paso.")
    r = WriteBullet(ws, r, "Datos_Proveedores", "Hoja de trabajo principal. Aqui se cargan los XMLs y se depuran las operaciones.")
    r = WriteBullet(ws, r, "Datos_Concentrados", "Se genera automaticamente al concentrar. Agrupa operaciones por RFC.")
    r = WriteBullet(ws, r, "DIOT2025", "Formato DIOT para declaraciones del ejercicio 2025 en adelante.")
    r = WriteBullet(ws, r, "DIOT2024", "Formato DIOT para declaraciones del ejercicio 2024 y anteriores.")
    r = r + 1

    ' =============================================
    ' FLUJO DE TRABAJO
    ' =============================================
    r = WriteHeader(ws, r, "Flujo de trabajo")
    r = r + 1

    ' Paso 1
    r = WriteStep(ws, r, "1", "Cargar XMLs")
    r = WriteText(ws, r, "Ve a la hoja ""Datos_Proveedores"" y haz clic en el boton ""Cargar XML"".")
    r = WriteText(ws, r, "Selecciona la carpeta que contiene los archivos XML de tus CFDIs.")
    r = WriteText(ws, r, "El macro lee cada XML y extrae: RFC, nombre, fecha, serie, folio, tipo, metodo de pago,")
    r = WriteText(ws, r, "desglose de IVA por tasa (16%, 8%, 0%, exento), retenciones y total.")
    r = r + 1
    r = WriteNote(ws, r, "Solo se procesan comprobantes tipo Ingreso (I) y Egreso (E).")
    r = WriteNote(ws, r, "Los egresos (notas de credito) se registran con valores negativos.")
    r = WriteNote(ws, r, "Si cargas la misma carpeta dos veces, los UUIDs duplicados se omiten automaticamente.")
    r = r + 1

    ' Paso 2
    r = WriteStep(ws, r, "2", "Depurar datos")
    r = WriteText(ws, r, "Revisa la informacion cargada en ""Datos_Proveedores"".")
    r = WriteText(ws, r, "Puedes agregar o eliminar filas manualmente si es necesario.")
    r = WriteText(ws, r, "Puedes cargar XMLs de otras carpetas; se agregaran sin duplicar UUIDs.")
    r = r + 1

    ' Paso 3
    r = WriteStep(ws, r, "3", "Concentrar por proveedor")
    r = WriteText(ws, r, "Haz clic en el boton ""Concentrar Proveedores"".")
    r = WriteText(ws, r, "Se creara la hoja ""Datos_Concentrados"" con los totales agrupados por RFC:")
    r = WriteText(ws, r, "nombre, numero de operaciones, bases gravadas, IVA trasladado, IVA retenido y total.")
    r = r + 1
    r = WriteNote(ws, r, "Las cantidades se redondean a enteros (sin decimales) como lo requiere el SAT.")
    r = r + 1

    ' Paso 4
    r = WriteStep(ws, r, "4", "Llenar formato DIOT")
    r = WriteText(ws, r, "Elige la hoja DIOT segun el ejercicio fiscal:")
    r = WriteBullet(ws, r, "DIOT2025", "Para declaraciones del ejercicio 2025 en adelante.")
    r = WriteBullet(ws, r, "DIOT2024", "Para declaraciones del ejercicio 2024 y anteriores.")
    r = r + 1
    r = WriteText(ws, r, "Llena las columnas conforme al formato del SAT, usando como referencia")
    r = WriteText(ws, r, "los datos concentrados. Los campos principales son:")
    r = r + 1
    r = WriteBullet(ws, r, "Tipo de tercero", "04 Nacional, 05 Extranjero, 15 Global")
    r = WriteBullet(ws, r, "Tipo de operacion", "03 Serv. Profesionales, 06 Arrendamiento, 85 Otros, etc.")
    r = WriteBullet(ws, r, "RFC", "Registro Federal de Contribuyentes del proveedor")
    r = WriteBullet(ws, r, "Valores de actos", "Montos sin decimales, separados por tasa de IVA")
    r = WriteBullet(ws, r, "IVA acreditable", "Montos de IVA pagado, sin decimales")
    r = r + 1
    r = WriteNote(ws, r, "Consulta el instructivo oficial del SAT para el detalle de cada campo.")
    r = WriteNote(ws, r, "Las cantidades NO deben llevar decimales ni signo de pesos.")
    r = r + 1

    ' Paso 5
    r = WriteStep(ws, r, "5", "Exportar archivo TXT")
    r = WriteText(ws, r, "Posicionate en la hoja DIOT que llenaste (DIOT2025 o DIOT2024).")
    r = WriteText(ws, r, "Haz clic en el boton ""Exportar DIOT"".")
    r = WriteText(ws, r, "Se generara un archivo TXT en la misma carpeta del libro con formato:")
    r = WriteText(ws, r, "DIOT_[NombreHoja]_CargaMasiva.txt")
    r = r + 1
    r = WriteNote(ws, r, "El archivo se genera en codificacion UTF-8 con BOM, listo para el portal del SAT.")
    r = WriteNote(ws, r, "Los campos se separan con el caracter pipe (|) como lo requiere el SAT.")
    r = r + 1

    ' =============================================
    ' BOTONES DISPONIBLES
    ' =============================================
    r = WriteHeader(ws, r, "Botones disponibles (hoja Datos_Proveedores)")
    r = r + 1

    Dim tblTop As Long
    tblTop = r

    ' Encabezados de tabla
    ws.Cells(r, 2).Value = "Boton"
    ws.Cells(r, 3).Value = "Accion"
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Font.Bold = True
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Font.Color = vbWhite
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Interior.Color = RGB(44, 62, 80)
    r = r + 1

    ' Filas de tabla
    r = WriteTableRow(ws, r, "Cargar XML", "Lee archivos XML de una carpeta y carga los datos en Datos_Proveedores.")
    r = WriteTableRow(ws, r, "Concentrar Proveedores", "Agrupa las operaciones por RFC y genera la hoja Datos_Concentrados.")
    r = WriteTableRow(ws, r, "Limpiar Datos", "Elimina todos los datos cargados y la hoja de concentrados (pide confirmacion).")
    r = WriteTableRow(ws, r, "Exportar DIOT", "Genera el archivo TXT para carga masiva (disponible en hojas DIOT2025/DIOT2024).")

    ' Bordes de tabla
    With ws.Range(ws.Cells(tblTop, 2), ws.Cells(r - 1, 3)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(189, 195, 199)
    End With
    ws.Columns("C").ColumnWidth = 75
    r = r + 1

    ' =============================================
    ' COLUMNAS DE DATOS_PROVEEDORES
    ' =============================================
    r = WriteHeader(ws, r, "Columnas de Datos_Proveedores")
    r = WriteText(ws, r, "Al cargar XMLs, cada comprobante se registra en una fila con las siguientes columnas:")
    r = r + 1

    tblTop = r
    ws.Cells(r, 2).Value = "Columna"
    ws.Cells(r, 3).Value = "Descripcion"
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Font.Bold = True
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Font.Color = vbWhite
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Interior.Color = RGB(44, 62, 80)
    r = r + 1

    r = WriteTableRow(ws, r, "A - Fecha", "Fecha de emision del comprobante (AAAA-MM-DD)")
    r = WriteTableRow(ws, r, "B - Serie", "Serie del comprobante (puede estar vacia)")
    r = WriteTableRow(ws, r, "C - Folio", "Numero de folio del comprobante")
    r = WriteTableRow(ws, r, "D - Tipo", "I = Ingreso (factura), E = Egreso (nota de credito)")
    r = WriteTableRow(ws, r, "E - RFC", "RFC del emisor (proveedor)")
    r = WriteTableRow(ws, r, "F - Nombre", "Nombre o razon social del emisor")
    r = WriteTableRow(ws, r, "G - Base Grav 16%", "Base gravada a tasa del 16% de IVA")
    r = WriteTableRow(ws, r, "H - Base Grav 8%", "Base gravada a tasa del 8% de IVA (region fronteriza)")
    r = WriteTableRow(ws, r, "I - Base Tasa 0%", "Base gravada a tasa 0% de IVA")
    r = WriteTableRow(ws, r, "J - Base Exento", "Base exenta de IVA")
    r = WriteTableRow(ws, r, "K - Descuento", "Descuento global del comprobante")
    r = WriteTableRow(ws, r, "L - IVA Trasl 16%", "IVA trasladado a tasa del 16%")
    r = WriteTableRow(ws, r, "M - IVA Trasl 8%", "IVA trasladado a tasa del 8%")
    r = WriteTableRow(ws, r, "N - IVA Retenido", "IVA retenido por el emisor")
    r = WriteTableRow(ws, r, "O - Total", "Monto total del comprobante")
    r = WriteTableRow(ws, r, "P - UUID", "Folio fiscal (Timbre Fiscal Digital)")
    r = WriteTableRow(ws, r, "Q - Metodo Pago", "PUE (pago en una exhibicion) o PPD (pago en parcialidades)")

    With ws.Range(ws.Cells(tblTop, 2), ws.Cells(r - 1, 3)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(189, 195, 199)
    End With
    r = r + 1

    ' =============================================
    ' NOTAS IMPORTANTES
    ' =============================================
    r = WriteHeader(ws, r, "Notas importantes")
    r = r + 1
    r = WriteNote(ws, r, "Habilitar macros: Al abrir el archivo, haz clic en ""Habilitar contenido"" si aparece la barra de seguridad.")
    r = WriteNote(ws, r, "Los egresos (notas de credito) se registran con valores negativos para que se neten automaticamente al concentrar.")
    r = WriteNote(ws, r, "El formato DIOT del SAT NO acepta decimales. La concentracion redondea al entero mas cercano.")
    r = WriteNote(ws, r, "El archivo TXT exportado usa codificacion UTF-8 y separador pipe (|) como lo requiere el portal del SAT.")
    r = WriteNote(ws, r, "La columna ""Pais"" en las hojas DIOT acepta el nombre del pais; al exportar se convierte automaticamente al codigo ISO.")
    r = WriteNote(ws, r, "Para mas informacion consulta el instructivo oficial: Instructivo_DIOT_V2_1102025.pdf")
    r = r + 2

    ' =============================================
    ' PIE
    ' =============================================
    With ws.Cells(r, 2)
        .Value = "Version 2.0 | Febrero 2025"
        .Font.Size = 9
        .Font.Color = RGB(149, 165, 166)
        .Font.Italic = True
    End With

    ' Proteger hoja (sin contrasena, solo para evitar edicion accidental)
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        AllowFiltering:=True

    ' Posicionar vista
    ws.Activate
    ws.Range("B2").Select
    ActiveWindow.DisplayHeadings = False

    MsgBox "Hoja de instrucciones generada correctamente.", vbInformation, "Instrucciones"
End Sub

' ============================================================
' FUNCIONES AUXILIARES DE FORMATO
' ============================================================

Private Function WriteHeader(ws As Worksheet, r As Long, texto As String) As Long
    With ws.Cells(r, 2)
        .Value = texto
        .Font.Size = 15
        .Font.Bold = True
        .Font.Color = RGB(44, 62, 80)
    End With
    WriteHeader = r + 1
End Function

Private Function WriteStep(ws As Worksheet, r As Long, numero As String, titulo As String) As Long
    With ws.Cells(r, 2)
        .Value = "Paso " & numero & ": " & titulo
        .Font.Size = 13
        .Font.Bold = True
        .Font.Color = RGB(52, 152, 219)
    End With
    WriteStep = r + 1
End Function

Private Function WriteText(ws As Worksheet, r As Long, texto As String) As Long
    With ws.Cells(r, 2)
        .Value = texto
        .Font.Size = 11
        .Font.Color = RGB(44, 62, 80)
    End With
    WriteText = r + 1
End Function

Private Function WriteBullet(ws As Worksheet, r As Long, etiqueta As String, descripcion As String) As Long
    With ws.Cells(r, 2)
        .Value = "     " & etiqueta
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(44, 62, 80)
    End With
    With ws.Cells(r, 3)
        .Value = descripcion
        .Font.Size = 11
        .Font.Color = RGB(100, 100, 100)
    End With
    WriteBullet = r + 1
End Function

Private Function WriteNote(ws As Worksheet, r As Long, texto As String) As Long
    With ws.Cells(r, 2)
        .Value = "  *  " & texto
        .Font.Size = 10
        .Font.Color = RGB(127, 140, 141)
        .Font.Italic = True
    End With
    WriteNote = r + 1
End Function

Private Function WriteTableRow(ws As Worksheet, r As Long, col1 As String, col2 As String) As Long
    ws.Cells(r, 2).Value = col1
    ws.Cells(r, 3).Value = col2

    ' Alternar color de fondo
    If r Mod 2 = 0 Then
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Interior.Color = RGB(245, 247, 249)
    End If

    WriteTableRow = r + 1
End Function
