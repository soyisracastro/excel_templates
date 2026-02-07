Attribute VB_Name = "ModuloXMLCFDI"
Option Explicit

' Namespaces CFDI 4.0
Private Const NS_CFDI As String = "xmlns:cfdi='http://www.sat.gob.mx/cfd/4'"
Private Const NS_TFD As String = "xmlns:tfd='http://www.sat.gob.mx/TimbreFiscalDigital'"

' Nombres de hojas
Private Const HOJA_PROVEEDORES As String = "Datos_Proveedores"
Private Const HOJA_CONCENTRADOS As String = "Datos_Concentrados"

' Layout de Datos_Proveedores
Private Const FILA_ENCABEZADOS As Long = 4
Private Const FILA_INICIO_DATOS As Long = 5

' Columnas de Datos_Proveedores
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

' ============================================================
' MACRO 1: Cargar XMLs de Ingresos y Egresos
' ============================================================
Sub CargarXMLProveedores()
    On Error GoTo ErrorHandler

    Dim folderPath As String
    Dim fso As Object
    Dim carpeta As Object
    Dim archivo As Object
    Dim xmlDoc As Object
    Dim wsProveedores As Worksheet
    Dim dicUUIDs As Object
    Dim nextRow As Long

    ' Contadores
    Dim contProcesados As Long, contDuplicados As Long, contIgnorados As Long
    contProcesados = 0: contDuplicados = 0: contIgnorados = 0

    ' Variables de extraccion
    Dim tipoComprobante As String, rfc As String, nombre As String, uuid As String
    Dim fecha As String, serie As String, folio As String, metodoPago As String
    Dim descuento As Double, total As Double
    Dim signo As Long

    ' Variables de desglose de impuestos
    Dim baseGrav16 As Double, baseGrav8 As Double, baseTasa0 As Double, baseExento As Double
    Dim ivaTrasl16 As Double, ivaTrasl8 As Double, ivaRetenido As Double
    Dim nodeList As Object, nodo As Object
    Dim impuesto As String, tipoFactor As String, tasaOCuota As String

    ' Variables para error handler
    Dim errNum As Long, errDesc As String

    ' 1. Seleccionar carpeta
    folderPath = SeleccionarCarpeta()
    If folderPath = "" Then Exit Sub

    ' Normalizar ruta
    folderPath = Trim(folderPath)
    If Right(folderPath, 1) = "\" Then folderPath = Left(folderPath, Len(folderPath) - 1)

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 2. Validar carpeta (doble check)
    Dim carpetaExiste As Boolean
    On Error Resume Next
    carpetaExiste = (Dir(folderPath, vbDirectory) <> "") Or fso.FolderExists(folderPath)
    On Error GoTo ErrorHandler

    If Not carpetaExiste Then
        Dim msg As String
        msg = "No se pudo encontrar la ruta." & vbCrLf & vbCrLf
        msg = msg & "Ruta detectada: " & folderPath & vbCrLf & vbCrLf

        If InStr(1, folderPath, "https://") > 0 Or InStr(1, folderPath, "sharepoint") > 0 Then
            msg = msg & "Estás intentando usar una ruta web de OneDrive/SharePoint." & vbCrLf
            msg = msg & "Abre la carpeta en el explorador, copia la ruta local y asegúrate de que los archivos estén 'Disponibles siempre en este dispositivo'."
        Else
            msg = msg & "Verifique que la carpeta no sea un acceso directo o una unidad de red desconectada."
        End If

        MsgBox msg, vbCritical, "Fallo al acceder a la carpeta"
        Exit Sub
    End If

    ' Acceder a la carpeta
    On Error Resume Next
    Set carpeta = fso.GetFolder(folderPath)
    If Err.Number <> 0 Then
        Dim errorNum As Long, errorDesc As String
        errorNum = Err.Number
        errorDesc = Err.Description
        On Error GoTo ErrorHandler

        MsgBox "Error " & errorNum & " al acceder a la carpeta:" & vbCrLf & vbCrLf & _
               "Ruta: " & folderPath & vbCrLf & _
               "Descripción: " & errorDesc & vbCrLf & vbCrLf & _
               "SOLUCIÓN: Los archivos de OneDrive pueden estar 'solo en la nube'." & vbCrLf & _
               "1. Abre la carpeta en el Explorador de Windows" & vbCrLf & _
               "2. Clic derecho > 'Mantener siempre en este dispositivo'" & vbCrLf & _
               "3. Espera a que se descarguen e intenta de nuevo.", _
               vbCritical, "Error de acceso a carpeta"
        Exit Sub
    End If
    On Error GoTo ErrorHandler

    ' 3. Verificar que existe la hoja Datos_Proveedores
    On Error Resume Next
    Set wsProveedores = ThisWorkbook.Sheets(HOJA_PROVEEDORES)
    On Error GoTo ErrorHandler

    If wsProveedores Is Nothing Then
        MsgBox "La hoja '" & HOJA_PROVEEDORES & "' no existe en este libro." & vbCrLf & _
               "Cree la hoja con los encabezados en la fila 4 antes de ejecutar este macro.", _
               vbCritical, "Hoja no encontrada"
        Exit Sub
    End If

    ' 4. Cargar UUIDs existentes para deduplicacion
    Set dicUUIDs = CargarUUIDsExistentes(wsProveedores)

    ' 5. Obtener siguiente fila disponible
    nextRow = ObtenerSiguienteFila(wsProveedores)

    ' 6. Configurar rendimiento
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Procesando XMLs de: " & folderPath

    ' 7. Iterar archivos XML
    For Each archivo In carpeta.Files
        If LCase(fso.GetExtensionName(archivo.Name)) = "xml" Then

            ' a. Parsear XML
            Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
            xmlDoc.async = False
            xmlDoc.Load archivo.Path

            If xmlDoc.parseError.ErrorCode <> 0 Then GoTo SiguienteArchivo

            ' b. Configurar namespaces
            xmlDoc.SetProperty "SelectionNamespaces", NS_CFDI & " " & NS_TFD

            ' c. Verificar tipo de comprobante
            tipoComprobante = GetAttr(xmlDoc, "/cfdi:Comprobante", "TipoDeComprobante")

            If tipoComprobante <> "I" And tipoComprobante <> "E" Then
                contIgnorados = contIgnorados + 1
                GoTo SiguienteArchivo
            End If

            ' d. Verificar UUID (dedup)
            uuid = UCase(Trim(GetAttr(xmlDoc, "/cfdi:Comprobante/cfdi:Complemento/tfd:TimbreFiscalDigital", "UUID")))

            If uuid = "" Then GoTo SiguienteArchivo

            If dicUUIDs.exists(uuid) Then
                contDuplicados = contDuplicados + 1
                GoTo SiguienteArchivo
            End If

            ' e. Extraer datos basicos
            rfc = GetAttr(xmlDoc, "/cfdi:Comprobante/cfdi:Emisor", "Rfc")
            nombre = GetAttr(xmlDoc, "/cfdi:Comprobante/cfdi:Emisor", "Nombre")
            fecha = Left(GetAttr(xmlDoc, "/cfdi:Comprobante", "Fecha"), 10)
            serie = GetAttr(xmlDoc, "/cfdi:Comprobante", "Serie")
            folio = GetAttr(xmlDoc, "/cfdi:Comprobante", "Folio")
            metodoPago = GetAttr(xmlDoc, "/cfdi:Comprobante", "MetodoPago")
            descuento = CDbl(Val(GetAttr(xmlDoc, "/cfdi:Comprobante", "Descuento")))
            total = CDbl(Val(GetAttr(xmlDoc, "/cfdi:Comprobante", "Total")))

            ' f. Determinar signo (Egresos = valores negativos)
            signo = IIf(tipoComprobante = "E", -1, 1)

            ' g. Inicializar acumuladores de impuestos
            baseGrav16 = 0: baseGrav8 = 0: baseTasa0 = 0: baseExento = 0
            ivaTrasl16 = 0: ivaTrasl8 = 0: ivaRetenido = 0

            ' h. Iterar nodos globales de Traslados
            Set nodeList = xmlDoc.SelectNodes("/cfdi:Comprobante/cfdi:Impuestos/cfdi:Traslados/cfdi:Traslado")

            If Not nodeList Is Nothing Then
                For Each nodo In nodeList
                    impuesto = GetNodeAttr(nodo, "Impuesto")

                    ' Solo IVA (002), ignorar IEPS (003) y otros
                    If impuesto = "002" Then
                        tipoFactor = GetNodeAttr(nodo, "TipoFactor")
                        tasaOCuota = GetNodeAttr(nodo, "TasaOCuota")

                        If tipoFactor = "Exento" Then
                            baseExento = baseExento + CDbl(Val(GetNodeAttr(nodo, "Base")))
                        ElseIf tasaOCuota = "0.160000" Then
                            baseGrav16 = baseGrav16 + CDbl(Val(GetNodeAttr(nodo, "Base")))
                            ivaTrasl16 = ivaTrasl16 + CDbl(Val(GetNodeAttr(nodo, "Importe")))
                        ElseIf tasaOCuota = "0.080000" Then
                            baseGrav8 = baseGrav8 + CDbl(Val(GetNodeAttr(nodo, "Base")))
                            ivaTrasl8 = ivaTrasl8 + CDbl(Val(GetNodeAttr(nodo, "Importe")))
                        ElseIf tasaOCuota = "0.000000" Then
                            baseTasa0 = baseTasa0 + CDbl(Val(GetNodeAttr(nodo, "Base")))
                        End If
                    End If
                Next nodo
            End If

            ' i. Iterar nodos globales de Retenciones (solo IVA)
            Set nodeList = xmlDoc.SelectNodes("/cfdi:Comprobante/cfdi:Impuestos/cfdi:Retenciones/cfdi:Retencion")

            If Not nodeList Is Nothing Then
                For Each nodo In nodeList
                    If GetNodeAttr(nodo, "Impuesto") = "002" Then
                        ivaRetenido = ivaRetenido + CDbl(Val(GetNodeAttr(nodo, "Importe")))
                    End If
                Next nodo
            End If

            ' j. Escribir fila en la hoja
            With wsProveedores
                .Cells(nextRow, COL_RFC).Value = rfc
                .Cells(nextRow, COL_NOMBRE).Value = nombre
                .Cells(nextRow, COL_UUID).Value = uuid
                .Cells(nextRow, COL_FECHA).Value = fecha
                .Cells(nextRow, COL_FOLIO).Value = Trim(serie & folio)
                .Cells(nextRow, COL_TIPO).Value = tipoComprobante
                .Cells(nextRow, COL_METODO).Value = metodoPago
                .Cells(nextRow, COL_GRAV16).Value = baseGrav16 * signo
                .Cells(nextRow, COL_GRAV8).Value = baseGrav8 * signo
                .Cells(nextRow, COL_TASA0).Value = baseTasa0 * signo
                .Cells(nextRow, COL_EXENTO).Value = baseExento * signo
                .Cells(nextRow, COL_DESCUENTO).Value = descuento * signo
                .Cells(nextRow, COL_IVA16).Value = ivaTrasl16 * signo
                .Cells(nextRow, COL_IVA8).Value = ivaTrasl8 * signo
                .Cells(nextRow, COL_IVARET).Value = ivaRetenido * signo
                .Cells(nextRow, COL_TOTAL).Value = total * signo
            End With

            ' k. Actualizar tracking
            dicUUIDs.Add uuid, 1
            nextRow = nextRow + 1
            contProcesados = contProcesados + 1

SiguienteArchivo:
        End If
    Next archivo

    ' 8. Mostrar resumen
    Dim msgResult As String
    msgResult = "Proceso completado." & vbCrLf & vbCrLf & _
        "XMLs cargados (Ingreso/Egreso): " & contProcesados & vbCrLf & _
        "Duplicados omitidos (UUID): " & contDuplicados & vbCrLf & _
        "Ignorados (Pagos/Nómina/Traslado): " & contIgnorados

Cleanup:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    If contProcesados > 0 Then
        MsgBox msgResult, vbInformation, "Carga de XML Completada"
    Else
        MsgBox msgResult, vbExclamation, "Carga de XML Completada"
    End If
    Exit Sub

ErrorHandler:
    errNum = Err.Number
    errDesc = Err.Description
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error inesperado: " & errNum & " - " & errDesc, _
        vbCritical, "Error en Carga de XML"
End Sub

' ============================================================
' MACRO 2: Concentrar datos por RFC
' ============================================================
Sub ConcentrarDatos()
    On Error GoTo ErrorConcentrar

    Dim wsProveedores As Worksheet
    Dim wsConcentrados As Worksheet
    Dim dicRFC As Object
    Dim rngDatos As Variant
    Dim ultimaFila As Long
    Dim numFilas As Long
    Dim i As Long
    Dim rfc As String
    Dim datos As Variant

    ' 1. Validar hoja origen
    On Error Resume Next
    Set wsProveedores = ThisWorkbook.Sheets(HOJA_PROVEEDORES)
    On Error GoTo ErrorConcentrar

    If wsProveedores Is Nothing Then
        MsgBox "La hoja '" & HOJA_PROVEEDORES & "' no existe.", vbCritical, "Error"
        Exit Sub
    End If

    ultimaFila = wsProveedores.Cells(wsProveedores.Rows.Count, COL_RFC).End(xlUp).Row

    If ultimaFila < FILA_INICIO_DATOS Then
        MsgBox "No hay datos para concentrar en '" & HOJA_PROVEEDORES & "'.", vbExclamation, "Sin datos"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "Concentrando datos por RFC..."

    ' 2. Leer datos a memoria
    rngDatos = wsProveedores.Range( _
        wsProveedores.Cells(FILA_INICIO_DATOS, 1), _
        wsProveedores.Cells(ultimaFila, ULTIMA_COL)).Value

    numFilas = UBound(rngDatos, 1)

    ' 3. Agrupar por RFC
    Set dicRFC = CreateObject("Scripting.Dictionary")

    For i = 1 To numFilas
        rfc = Trim(CStr(rngDatos(i, COL_RFC)))
        If rfc = "" Then GoTo SiguienteFilaConc

        If Not dicRFC.exists(rfc) Then
            ' Array: (0)Nombre, (1)NumOps, (2)Grav16, (3)Grav8, (4)Tasa0,
            '        (5)Exento, (6)Descuento, (7)IVA16, (8)IVA8, (9)IVARet, (10)Total
            dicRFC.Add rfc, Array( _
                CStr(rngDatos(i, COL_NOMBRE)), _
                1, _
                CDbl(Val(rngDatos(i, COL_GRAV16))), _
                CDbl(Val(rngDatos(i, COL_GRAV8))), _
                CDbl(Val(rngDatos(i, COL_TASA0))), _
                CDbl(Val(rngDatos(i, COL_EXENTO))), _
                CDbl(Val(rngDatos(i, COL_DESCUENTO))), _
                CDbl(Val(rngDatos(i, COL_IVA16))), _
                CDbl(Val(rngDatos(i, COL_IVA8))), _
                CDbl(Val(rngDatos(i, COL_IVARET))), _
                CDbl(Val(rngDatos(i, COL_TOTAL))) _
            )
        Else
            datos = dicRFC(rfc)
            datos(1) = datos(1) + 1
            datos(2) = datos(2) + CDbl(Val(rngDatos(i, COL_GRAV16)))
            datos(3) = datos(3) + CDbl(Val(rngDatos(i, COL_GRAV8)))
            datos(4) = datos(4) + CDbl(Val(rngDatos(i, COL_TASA0)))
            datos(5) = datos(5) + CDbl(Val(rngDatos(i, COL_EXENTO)))
            datos(6) = datos(6) + CDbl(Val(rngDatos(i, COL_DESCUENTO)))
            datos(7) = datos(7) + CDbl(Val(rngDatos(i, COL_IVA16)))
            datos(8) = datos(8) + CDbl(Val(rngDatos(i, COL_IVA8)))
            datos(9) = datos(9) + CDbl(Val(rngDatos(i, COL_IVARET)))
            datos(10) = datos(10) + CDbl(Val(rngDatos(i, COL_TOTAL)))
            dicRFC(rfc) = datos
        End If
SiguienteFilaConc:
    Next i

    ' 4. Crear/recrear hoja de concentrados
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(HOJA_CONCENTRADOS).Delete
    On Error GoTo ErrorConcentrar
    Application.DisplayAlerts = True

    Set wsConcentrados = ThisWorkbook.Sheets.Add(After:=wsProveedores)
    wsConcentrados.Name = HOJA_CONCENTRADOS

    ' 5. Escribir encabezados
    Dim headers As Variant
    headers = Array("RFC", "Nombre del Emisor", "Num. Operaciones", _
        "Valor Actos Gravados 16%", "Valor Actos Gravados 8%", _
        "Valor Actos Tasa 0%", "Valor Actos Exentos", "Descuento", _
        "IVA Trasladado 16%", "IVA Trasladado 8%", "IVA Retenido", "Total")

    wsConcentrados.Range("A1:L1").Value = headers

    With wsConcentrados.Range("A1:L1")
        .Font.Bold = True
        .Interior.Color = RGB(44, 62, 80)
        .Font.Color = vbWhite
    End With

    ' 6. Escribir datos
    Dim fila As Long
    fila = 2
    Dim key As Variant

    For Each key In dicRFC.Keys
        datos = dicRFC(key)
        wsConcentrados.Cells(fila, 1).Value = key         ' RFC
        wsConcentrados.Cells(fila, 2).Value = datos(0)    ' Nombre
        wsConcentrados.Cells(fila, 3).Value = datos(1)    ' Num. Operaciones
        wsConcentrados.Cells(fila, 4).Value = datos(2)    ' Grav 16%
        wsConcentrados.Cells(fila, 5).Value = datos(3)    ' Grav 8%
        wsConcentrados.Cells(fila, 6).Value = datos(4)    ' Tasa 0%
        wsConcentrados.Cells(fila, 7).Value = datos(5)    ' Exentos
        wsConcentrados.Cells(fila, 8).Value = datos(6)    ' Descuento
        wsConcentrados.Cells(fila, 9).Value = datos(7)    ' IVA 16%
        wsConcentrados.Cells(fila, 10).Value = datos(8)   ' IVA 8%
        wsConcentrados.Cells(fila, 11).Value = datos(9)   ' IVA Ret
        wsConcentrados.Cells(fila, 12).Value = datos(10)  ' Total
        fila = fila + 1
    Next key

    ' 7. Formato
    wsConcentrados.Columns("A:L").AutoFit

    If fila > 2 Then
        wsConcentrados.Range("D2:L" & fila - 1).NumberFormat = "$#,##0.00"

        With wsConcentrados.Range("A1:L" & fila - 1).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(200, 200, 200)
        End With
    End If

    Application.StatusBar = False
    Application.ScreenUpdating = True

    MsgBox "Datos concentrados generados." & vbCrLf & _
        dicRFC.Count & " proveedores encontrados.", _
        vbInformation, "Concentración Completada"
    Exit Sub

ErrorConcentrar:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error al concentrar datos: " & Err.Number & " - " & Err.Description, _
        vbCritical, "Error"
End Sub

' ============================================================
' MACRO 3: Limpiar datos
' ============================================================
Sub LimpiarDatos()
    Dim resp As VbMsgBoxResult
    resp = MsgBox("¿Desea limpiar todos los datos cargados en '" & HOJA_PROVEEDORES & "'?" & vbCrLf & vbCrLf & _
        "Esta acción eliminará los datos desde la fila 5 y la hoja '" & HOJA_CONCENTRADOS & "' si existe." & vbCrLf & _
        "Esta acción no se puede deshacer.", _
        vbQuestion + vbYesNo, "Confirmar Limpieza")

    If resp = vbNo Then Exit Sub

    ' Limpiar Datos_Proveedores
    Dim wsProveedores As Worksheet
    On Error Resume Next
    Set wsProveedores = ThisWorkbook.Sheets(HOJA_PROVEEDORES)
    On Error GoTo 0

    If Not wsProveedores Is Nothing Then
        Dim ultimaFila As Long
        ultimaFila = wsProveedores.Cells(wsProveedores.Rows.Count, COL_RFC).End(xlUp).Row

        If ultimaFila >= FILA_INICIO_DATOS Then
            wsProveedores.Range( _
                wsProveedores.Cells(FILA_INICIO_DATOS, 1), _
                wsProveedores.Cells(ultimaFila, ULTIMA_COL)).ClearContents
        End If
    End If

    ' Eliminar Datos_Concentrados si existe
    Dim wsConc As Worksheet
    On Error Resume Next
    Set wsConc = ThisWorkbook.Sheets(HOJA_CONCENTRADOS)
    On Error GoTo 0

    If Not wsConc Is Nothing Then
        Application.DisplayAlerts = False
        wsConc.Delete
        Application.DisplayAlerts = True
    End If

    ' Activar hoja de proveedores
    If Not wsProveedores Is Nothing Then
        wsProveedores.Activate
        wsProveedores.Cells(FILA_INICIO_DATOS, COL_RFC).Select
    End If

    MsgBox "Datos limpiados correctamente." & vbCrLf & _
        "La hoja está lista para una nueva carga.", _
        vbInformation, "Limpieza Completada"
End Sub

' ============================================================
' FUNCIONES AUXILIARES
' ============================================================

Function SeleccionarCarpeta() As String
    Dim fd As Object
    Set fd = Application.FileDialog(4) ' msoFileDialogFolderPicker
    fd.Title = "Selecciona la carpeta que contiene los archivos XML"
    If fd.Show = -1 Then
        SeleccionarCarpeta = fd.SelectedItems(1)
    Else
        SeleccionarCarpeta = ""
    End If
End Function

Private Function GetAttr(xmlDoc As Object, xpath As String, attrName As String) As String
    Dim node As Object
    Set node = xmlDoc.SelectSingleNode(xpath)
    If Not node Is Nothing Then
        If Not node.Attributes.getNamedItem(attrName) Is Nothing Then
            GetAttr = node.Attributes.getNamedItem(attrName).Text
        End If
    End If
End Function

Private Function GetNodeAttr(node As Object, attrName As String) As String
    If Not node Is Nothing Then
        If Not node.Attributes.getNamedItem(attrName) Is Nothing Then
            GetNodeAttr = node.Attributes.getNamedItem(attrName).Text
        End If
    End If
End Function

Private Function CargarUUIDsExistentes(ws As Worksheet) As Object
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")

    Dim ultimaFila As Long
    ultimaFila = ws.Cells(ws.Rows.Count, COL_UUID).End(xlUp).Row

    If ultimaFila < FILA_INICIO_DATOS Then
        Set CargarUUIDsExistentes = dic
        Exit Function
    End If

    ' Leer columna UUID completa a array para velocidad
    Dim rng As Variant
    rng = ws.Range(ws.Cells(FILA_INICIO_DATOS, COL_UUID), _
                   ws.Cells(ultimaFila, COL_UUID)).Value

    Dim i As Long
    Dim uuid As String
    For i = 1 To UBound(rng, 1)
        uuid = UCase(Trim(CStr(rng(i, 1))))
        If uuid <> "" And Not dic.exists(uuid) Then
            dic.Add uuid, 1
        End If
    Next i

    Set CargarUUIDsExistentes = dic
End Function

Private Function ObtenerSiguienteFila(ws As Worksheet) As Long
    Dim ultimaFila As Long
    ultimaFila = ws.Cells(ws.Rows.Count, COL_RFC).End(xlUp).Row

    If ultimaFila < FILA_INICIO_DATOS Then
        ObtenerSiguienteFila = FILA_INICIO_DATOS
    Else
        ObtenerSiguienteFila = ultimaFila + 1
    End If
End Function
