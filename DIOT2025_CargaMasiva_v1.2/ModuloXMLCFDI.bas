Attribute VB_Name = "ModuloXMLCFDI"
Option Explicit

' Constantes para Namespaces de CFDI 4.0
Private Const NS_CFDI As String = "xmlns:cfdi='http://www.sat.gob.mx/cfd/4'"
Private Const NS_TFD As String = "xmlns:tfd='http://www.sat.gob.mx/TimbreFiscalDigital'"
Private Const NS_PAGO20 As String = "xmlns:pago20='http://www.sat.gob.mx/Pagos20'"

Sub CargarXMLs()
    Dim folderPath As String
    Dim fileName As String
    Dim xmlDoc As Object
    Dim dicConsolidado As Object
    Dim fso As Object
    Dim carpeta As Object
    Dim archivo As Object
    Dim rfc As String, nombre As String, uuid As String
    Dim total As Double, subtotal As Double, ivaTrasladado As Double, ivaRetenido As Double
    Dim tipoComprobante As String, metodoPago As String
    
    ' Seleccionar carpeta
    folderPath = SeleccionarCarpeta()
    If folderPath = "" Then Exit Sub
    
    ' Normalizar ruta (limpiar espacios y barras finales)
    folderPath = Trim(folderPath)
    If Right(folderPath, 1) = "\" Then folderPath = Left(folderPath, Len(folderPath) - 1)
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Validación robusta: Intentar con Dir (VBA Nativo) y FSO (Windows Scripting)
    Dim carpetaExiste As Boolean
    On Error Resume Next
    carpetaExiste = (Dir(folderPath, vbDirectory) <> "") Or fso.FolderExists(folderPath)
    On Error GoTo 0
    
    If Not carpetaExiste Then
        Dim msg As String
        msg = "Error 76: No se pudo encontrar la ruta." & vbCrLf & vbCrLf
        msg = msg & "Ruta detectada: " & folderPath & vbCrLf & vbCrLf
        
        If InStr(1, folderPath, "https://") > 0 Or InStr(1, folderPath, "sharepoint") > 0 Then
            msg = msg & "ADVERTENCIA: Estás intentando usar una ruta web de OneDrive/SharePoint." & vbCrLf
            msg = msg & "Sugerencia: Abre la carpeta en tu explorador de archivos, copia la ruta real del disco local y asegúrate de que los archivos estén 'Disponibles siempre en este dispositivo'."
        Else
            msg = msg & "Sugerencia: Verifique que la carpeta no sea un acceso directo o una unidad de red desconectada."
        End If
        
        MsgBox msg, vbCritical, "Fallo al acceder a la carpeta"
        Exit Sub
    End If
    
    Set dicConsolidado = CreateObject("Scripting.Dictionary")
    
    ' Intentar acceder a la carpeta con manejo de errores
    On Error Resume Next
    Set carpeta = fso.GetFolder(folderPath)
    If Err.Number <> 0 Then
        Dim errorNum As Long, errorDesc As String
        errorNum = Err.Number
        errorDesc = Err.Description
        On Error GoTo 0
        
        MsgBox "Error " & errorNum & " al acceder a la carpeta:" & vbCrLf & vbCrLf & _
               "Ruta: " & folderPath & vbCrLf & _
               "Descripción: " & errorDesc & vbCrLf & vbCrLf & _
               "SOLUCIÓN: Los archivos de OneDrive pueden estar 'solo en la nube'." & vbCrLf & _
               "1. Abre la carpeta 'xml' en el Explorador de Windows" & vbCrLf & _
               "2. Clic derecho > 'Mantener siempre en este dispositivo'" & vbCrLf & _
               "3. Espera a que se descarguen los archivos e intenta de nuevo.", _
               vbCritical, "Error de acceso a carpeta OneDrive"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Mostrar barra de estado
    Application.StatusBar = "Procesando XMLs..."
    Application.ScreenUpdating = False
    
    For Each archivo In carpeta.Files
        If LCase(fso.GetExtensionName(archivo.Name)) = "xml" Then
            Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
            xmlDoc.async = False
            xmlDoc.Load archivo.Path
            
            If xmlDoc.parseError.ErrorCode = 0 Then
                ' Configurar Namespaces
                xmlDoc.SetProperty "SelectionNamespaces", NS_CFDI & " " & NS_TFD & " " & NS_PAGO20
                
                ' Obtener datos básicos
                tipoComprobante = GetAttr(xmlDoc, "/cfdi:Comprobante", "TipoDeComprobante")
                rfc = GetAttr(xmlDoc, "/cfdi:Comprobante/cfdi:Emisor", "Rfc")
                nombre = GetAttr(xmlDoc, "/cfdi:Comprobante/cfdi:Emisor", "Nombre")
                
                If tipoComprobante = "I" Then
                    ' Procesar Ingreso
                    subtotal = CDbl(Val(GetAttr(xmlDoc, "/cfdi:Comprobante", "SubTotal")))
                    total = CDbl(Val(GetAttr(xmlDoc, "/cfdi:Comprobante", "Total")))
                    metodoPago = GetAttr(xmlDoc, "/cfdi:Comprobante", "MetodoPago")
                    
                    ' Impuestos Globales
                    ivaTrasladado = CDbl(Val(GetAttr(xmlDoc, "/cfdi:Comprobante/cfdi:Impuestos", "TotalImpuestosTrasladados")))
                    ivaRetenido = CDbl(Val(GetAttr(xmlDoc, "/cfdi:Comprobante/cfdi:Impuestos", "TotalImpuestosRetenidos")))
                    
                    uuid = GetAttr(xmlDoc, "/cfdi:Comprobante/cfdi:Complemento/tfd:TimbreFiscalDigital", "UUID")
                    
                    ActualizarDiccionario dicConsolidado, rfc, nombre, subtotal, ivaTrasladado, ivaRetenido, total, uuid, metodoPago, "Ingreso"
                    
                ElseIf tipoComprobante = "P" Then
                    ' Procesar Pago (Complemento 2.0)
                    Dim pagos As Object, pago As Object, docRel As Object
                    Set pagos = xmlDoc.SelectNodes("//pago20:Pago")
                    
                    For Each pago In pagos
                        Dim montoPago As Double
                        montoPago = CDbl(Val(GetNodeAttr(pago, "Monto")))
                        
                        Set docRel = pago.SelectSingleNode("pago20:DoctoRelacionado")
                        If Not docRel Is Nothing Then
                            ' Extraer IVA del pago desde DoctoRelacionado o ImpuestosP
                            ' En Pagos 2.0 el IVA suele venir en ImpuestosP o ImpuestosDR
                            Dim ivaP As Double, baseP As Double
                            ivaP = CDbl(Val(GetAttr(xmlDoc, "//pago20:TrasladoP", "ImporteP")))
                            If ivaP = 0 Then ivaP = CDbl(Val(GetAttr(xmlDoc, "//pago20:TrasladoDR", "ImporteDR")))
                            
                            baseP = CDbl(Val(GetAttr(xmlDoc, "//pago20:TrasladoP", "BaseP")))
                            If baseP = 0 Then baseP = CDbl(Val(GetAttr(xmlDoc, "//pago20:TrasladoDR", "BaseDR")))
                            
                            uuid = GetNodeAttr(docRel, "IdDocumento")
                            ActualizarDiccionario dicConsolidado, rfc, nombre, baseP, ivaP, 0, montoPago, uuid, "PPD", "Pago"
                        End If
                    Next
                End If
            End If
        End If
    Next
    
    EscribirEnHoja dicConsolidado
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Proceso completado. Se han consolidado los datos en la hoja 'CFDI_Importados'.", vbInformation
End Sub

Function SeleccionarCarpeta() As String
    Dim fd As Object
    Set fd = Application.FileDialog(4) ' msoFileDialogFolderPicker = 4 (NO es 3)
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

Sub ActualizarDiccionario(dic As Object, rfc As String, nombre As String, subtotal As Double, ivaTras As Double, ivaRet As Double, total As Double, uuid As String, metodo As String, tipo As String)
    Dim key As String
    key = rfc
    
    If Not dic.exists(key) Then
        ' Array: [Nombre, Subtotal, IVA Tras, IVA Ret, Total, # Ops, UUIDs, Metodo]
        dic.Add key, Array(nombre, subtotal, ivaTras, ivaRet, total, 1, uuid, metodo)
    Else
        Dim datos As Variant
        datos = dic(key)
        datos(1) = datos(1) + subtotal
        datos(2) = datos(2) + ivaTras
        datos(3) = datos(3) + ivaRet
        datos(4) = datos(4) + total
        datos(5) = datos(5) + 1
        ' Evitar duplicar UUIDs si son iguales (poco probable en consolidación, pero por si acaso)
        If InStr(datos(6), uuid) = 0 Then datos(6) = datos(6) & ", " & uuid
        dic(key) = datos
    End If
End Sub

Sub EscribirEnHoja(dic As Object)
    Dim ws As Worksheet
    Dim nombreHoja As String
    nombreHoja = "CFDI_Importados"
    
    ' Eliminar hoja si existe
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets(nombreHoja).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set ws = Sheets.Add
    ws.Name = nombreHoja
    
    ' Encabezados con diseño premium (simulado por formato)
    Dim headers As Variant
    headers = Array("RFC", "Nombre", "Subtotal Acum.", "IVA Trasladado", "IVA Retenido", "Total Acum.", "Num. Facturas", "UUIDs Relacionados", "Método Pago Predominante")
    
    ws.Range("A1:I1").Value = headers
    ws.Range("A1:I1").Font.Bold = True
    ws.Range("A1:I1").Interior.Color = RGB(44, 62, 80)
    ws.Range("A1:I1").Font.Color = vbWhite
    
    Dim fila As Long
    fila = 2
    Dim key As Variant
    For Each key In dic.Keys
        Dim datos As Variant
        datos = dic(key)
        ws.Cells(fila, 1).Value = key
        ws.Cells(fila, 2).Value = datos(0)
        ws.Cells(fila, 3).Value = datos(1)
        ws.Cells(fila, 4).Value = datos(2)
        ws.Cells(fila, 5).Value = datos(3)
        ws.Cells(fila, 6).Value = datos(4)
        ws.Cells(fila, 7).Value = datos(5)
        ws.Cells(fila, 8).Value = datos(6)
        ws.Cells(fila, 9).Value = datos(7)
        fila = fila + 1
    Next
    
    ' Formato
    ws.Columns("A:I").AutoFit
    ws.Range("C2:F" & fila).NumberFormat = "$#,##0.00"
    
    ' Aplicar bordes suaves
    With ws.Range("A1:I" & fila - 1).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(200, 200, 200)
    End With
End Sub
