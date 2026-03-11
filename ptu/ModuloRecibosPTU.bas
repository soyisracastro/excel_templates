Attribute VB_Name = "ModuloRecibosPTU"
'==============================================================================
' ModuloRecibosPTU - Generador de Recibos de PTU en PDF
'==============================================================================
' Marca: Columna 13
'
' INSTRUCCIONES DE USO:
' 1. Guarda una copia del archivo .xlsx como .xlsm (con macros)
' 2. Abre el Editor de VBA (Alt+F11)
' 3. Archivo > Importar archivo > selecciona este .bas
' 4. Cierra el editor y ejecuta la macro "GenerarRecibosPTU"
'
' La macro genera recibos individuales en PDF para cada empleado
' y los guarda en una subcarpeta "Recibos_PTU_{ejercicio}/"
'==============================================================================

Option Explicit

Private Const DATOS_SHEET As String = "Datos"
Private Const CALCULO_SHEET As String = "C" & ChrW(225) & "lculo_ISR"
Private Const HEADER_ROW As Long = 13
Private Const DATA_START As Long = 14

Sub GenerarRecibosPTU()
    Dim wsDatos As Worksheet
    Dim wsCalc As Worksheet
    Dim ejercicio As Long
    Dim empresa As String
    Dim rfcEmpresa As String
    Dim fechaPago As Date
    Dim lastRow As Long
    Dim outputPath As String
    Dim opcion As VbMsgBoxResult
    Dim empleadoIdx As Long
    Dim i As Long
    Dim generados As Long

    On Error GoTo ErrorHandler

    ' Validar hojas
    Set wsDatos = ThisWorkbook.Sheets(DATOS_SHEET)
    Set wsCalc = ThisWorkbook.Sheets(CALCULO_SHEET)

    ' Leer datos empresa
    empresa = wsDatos.Range("B3").Value
    rfcEmpresa = wsDatos.Range("B4").Value
    ejercicio = wsDatos.Range("B5").Value
    fechaPago = wsDatos.Range("B3").Value  ' Config!FechaPago via named range

    If empresa = "" Then
        MsgBox "No se ha capturado el nombre de la empresa en la hoja Datos.", _
               vbExclamation, "Datos incompletos"
        Exit Sub
    End If

    ' Encontrar ultima fila con datos
    lastRow = DATA_START
    Do While wsDatos.Cells(lastRow, 2).Value <> "" And lastRow <= 63
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1

    If lastRow < DATA_START Then
        MsgBox "No hay empleados registrados en la hoja Datos.", _
               vbExclamation, "Sin empleados"
        Exit Sub
    End If

    ' Preguntar: todos o uno
    opcion = MsgBox("" & ChrW(191) & "Generar recibos para TODOS los empleados?" & vbCrLf & vbCrLf & _
                    "S" & ChrW(237) & " = Todos los empleados" & vbCrLf & _
                    "No = Seleccionar uno", _
                    vbYesNoCancel + vbQuestion, "Generar Recibos PTU")

    If opcion = vbCancel Then Exit Sub

    ' Si es uno solo, pedir numero
    If opcion = vbNo Then
        Dim input_str As String
        input_str = InputBox("Escribe el n" & ChrW(250) & "mero de fila del empleado " & _
                             "(fila " & DATA_START & " a " & lastRow & "):", _
                             "Seleccionar empleado", CStr(DATA_START))
        If input_str = "" Then Exit Sub
        empleadoIdx = CLng(input_str)
        If empleadoIdx < DATA_START Or empleadoIdx > lastRow Then
            MsgBox "Fila fuera de rango.", vbExclamation
            Exit Sub
        End If
    End If

    ' Crear carpeta de salida
    outputPath = ThisWorkbook.Path & "\Recibos_PTU_" & ejercicio & "\"
    If Dir(outputPath, vbDirectory) = "" Then
        MkDir outputPath
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    generados = 0

    ' Generar recibos
    If opcion = vbYes Then
        ' Todos
        For i = DATA_START To lastRow
            If wsDatos.Cells(i, 2).Value <> "" Then
                GenerarReciboIndividual wsDatos, wsCalc, i, empresa, rfcEmpresa, ejercicio, outputPath
                generados = generados + 1
            End If
        Next i
    Else
        ' Solo uno
        GenerarReciboIndividual wsDatos, wsCalc, empleadoIdx, empresa, rfcEmpresa, ejercicio, outputPath
        generados = 1
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Se generaron " & generados & " recibo(s) en:" & vbCrLf & outputPath, _
           vbInformation, "Recibos generados"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
End Sub

Private Sub GenerarReciboIndividual(wsDatos As Worksheet, wsCalc As Worksheet, _
                                     fila As Long, empresa As String, _
                                     rfcEmpresa As String, ejercicio As Long, _
                                     outputPath As String)
    Dim wsTemp As Worksheet
    Dim nombre As String
    Dim rfc As String
    Dim curp As String
    Dim ptuBruta As Double
    Dim montoMax As Double
    Dim ptuReal As Double
    Dim ptuExenta As Double
    Dim ptuGravada As Double
    Dim isrRetenido As Double
    Dim metodoISR As String
    Dim ptuNeta As Double
    Dim irFila As Long
    Dim fileName As String

    ' Leer datos del empleado
    nombre = wsDatos.Cells(fila, 2).Value
    rfc = wsDatos.Cells(fila, 3).Value
    curp = wsDatos.Cells(fila, 4).Value
    ptuBruta = Val(CStr(wsDatos.Cells(fila, 20).Value))  ' Col T
    montoMax = Val(CStr(wsDatos.Cells(fila, 23).Value))   ' Col W
    ptuReal = Val(CStr(wsDatos.Cells(fila, 24).Value))    ' Col X

    irFila = 2 + (fila - DATA_START)  ' Calculo_ISR row
    ptuExenta = Val(CStr(wsCalc.Cells(irFila, 4).Value))   ' Col D
    ptuGravada = Val(CStr(wsCalc.Cells(irFila, 5).Value))  ' Col E
    isrRetenido = Val(CStr(wsCalc.Cells(irFila, 21).Value)) ' Col U
    metodoISR = CStr(wsCalc.Cells(irFila, 20).Value)        ' Col T
    ptuNeta = Val(CStr(wsCalc.Cells(irFila, 22).Value))     ' Col V

    ' Crear hoja temporal
    Set wsTemp = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsTemp.Name = "TmpRecibo"

    ' Formato pagina
    With wsTemp.PageSetup
        .PaperSize = xlPaperLetter
        .Orientation = xlPortrait
        .LeftMargin = Application.InchesToPoints(0.75)
        .RightMargin = Application.InchesToPoints(0.75)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .PrintArea = "A1:E30"
    End With

    wsTemp.Columns("A").ColumnWidth = 4
    wsTemp.Columns("B").ColumnWidth = 25
    wsTemp.Columns("C").ColumnWidth = 20
    wsTemp.Columns("D").ColumnWidth = 20
    wsTemp.Columns("E").ColumnWidth = 4

    ' --- Escribir recibo ---

    ' Empresa
    wsTemp.Range("B2:D2").Merge
    With wsTemp.Range("B2")
        .Value = empresa
        .Font.Name = "Aptos"
        .Font.Size = 16
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    wsTemp.Range("B3:D3").Merge
    With wsTemp.Range("B3")
        .Value = "RFC: " & rfcEmpresa
        .Font.Name = "Aptos"
        .Font.Size = 10
        .Font.Color = RGB(127, 140, 141)
        .HorizontalAlignment = xlCenter
    End With

    ' Titulo
    wsTemp.Range("B5:D5").Merge
    With wsTemp.Range("B5")
        .Value = "RECIBO DE PTU " & ChrW(8212) & " Ejercicio " & ejercicio
        .Font.Name = "Aptos"
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ' Datos empleado
    wsTemp.Range("B7").Value = "Trabajador:"
    wsTemp.Range("C7").Value = nombre
    wsTemp.Range("B8").Value = "RFC:"
    wsTemp.Range("C8").Value = rfc
    wsTemp.Range("B9").Value = "CURP:"
    wsTemp.Range("C9").Value = curp

    Dim rr As Long
    For rr = 7 To 9
        wsTemp.Cells(rr, 2).Font.Bold = True
        wsTemp.Cells(rr, 2).Font.Name = "Aptos"
        wsTemp.Cells(rr, 3).Font.Name = "Aptos"
    Next rr

    ' Tabla desglose
    wsTemp.Range("B11").Value = "Concepto"
    wsTemp.Range("C11").Value = "Importe"
    With wsTemp.Range("B11:C11")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Font.Name = "Aptos"
        .Interior.Color = RGB(44, 62, 80)
        .HorizontalAlignment = xlCenter
    End With

    Dim items() As Variant
    items = Array( _
        Array("PTU Bruta", ptuBruta), _
        Array("Tope aplicado", montoMax), _
        Array("PTU Real", ptuReal), _
        Array("PTU Exenta", ptuExenta), _
        Array("PTU Gravada", ptuGravada), _
        Array("ISR Retenido (" & metodoISR & ")", isrRetenido) _
    )

    Dim idx As Long
    For idx = 0 To UBound(items)
        rr = 12 + idx
        wsTemp.Cells(rr, 2).Value = items(idx)(0)
        wsTemp.Cells(rr, 3).Value = items(idx)(1)
        wsTemp.Cells(rr, 3).NumberFormat = "#,##0.00"
        wsTemp.Cells(rr, 2).Font.Name = "Aptos"
        wsTemp.Cells(rr, 3).Font.Name = "Aptos"
        If rr Mod 2 = 0 Then
            wsTemp.Range("B" & rr & ":C" & rr).Interior.Color = RGB(244, 246, 248)
        End If
    Next idx

    ' Neto destacado
    rr = 19
    wsTemp.Cells(rr, 2).Value = "PTU NETA A RECIBIR"
    wsTemp.Cells(rr, 3).Value = ptuNeta
    wsTemp.Cells(rr, 3).NumberFormat = "#,##0.00"
    With wsTemp.Range("B" & rr & ":C" & rr)
        .Font.Bold = True
        .Font.Size = 14
        .Font.Name = "Aptos"
        .Font.Color = RGB(30, 107, 58)
        .Interior.Color = RGB(213, 232, 212)
    End With

    ' Firma
    wsTemp.Range("B22:D22").Merge
    With wsTemp.Range("B22")
        .Value = "Recib" & ChrW(237) & " de conformidad la cantidad arriba se" & ChrW(241) & "alada."
        .Font.Name = "Aptos"
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
    End With

    wsTemp.Range("B25:D25").Merge
    With wsTemp.Range("B25")
        .Value = String(50, "_")
        .HorizontalAlignment = xlCenter
    End With

    wsTemp.Range("B26:D26").Merge
    With wsTemp.Range("B26")
        .Value = "Nombre y firma del trabajador"
        .Font.Name = "Aptos"
        .Font.Size = 10
        .Font.Color = RGB(127, 140, 141)
        .HorizontalAlignment = xlCenter
    End With

    wsTemp.Range("B28:D28").Merge
    With wsTemp.Range("B28")
        .Value = "De conformidad con los art" & ChrW(237) & "culos 117 al 131 de la Ley Federal del Trabajo."
        .Font.Name = "Aptos"
        .Font.Size = 8
        .Font.Italic = True
        .Font.Color = RGB(127, 140, 141)
        .HorizontalAlignment = xlCenter
    End With

    ' Exportar PDF
    fileName = outputPath & "Recibo_PTU_" & Replace(nombre, " ", "_") & ".pdf"
    wsTemp.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False

    ' Eliminar hoja temporal
    Application.DisplayAlerts = False
    wsTemp.Delete
    Application.DisplayAlerts = True
End Sub
