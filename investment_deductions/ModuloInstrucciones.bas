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
    ws.Columns("B").ColumnWidth = 120
    ActiveWindow.DisplayGridlines = False

    r = 2

    ' =============================================
    ' TITULO PRINCIPAL
    ' =============================================
    With ws.Cells(r, 2)
        .Value = "Deducci" & ChrW(243) & "n de Inversiones LISR"
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = RGB(44, 62, 80)
    End With
    r = r + 1

    With ws.Cells(r, 2)
        .Value = "Plantilla de C" & ChrW(225) & "lculo Fiscal - Art" & ChrW(237) & "culos 31 al 38"
        .Font.Size = 13
        .Font.Color = RGB(127, 140, 141)
    End With
    r = r + 2

    ' L" & ChrW(237) & "nea separadora
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 2))
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(52, 152, 219)
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    r = r + 2

    ' =============================================
    ' DESCRIPCION
    ' =============================================
    r = WriteHeader(ws, r, "Descripci" & ChrW(243) & "n")
    r = WriteText(ws, r, "Esta plantilla calcula autom" & ChrW(225) & "ticamente la deducci" & ChrW(243) & "n fiscal de inversiones conforme a la LISR.")
    r = WriteText(ws, r, "Incluye: cat" & ChrW(225) & "logo de porcentajes (Art. 33, 34 y 35), topes de autom" & ChrW(243) & "viles (Art. 36),")
    r = WriteText(ws, r, "actualizaci" & ChrW(243) & "n por INPC, y calculadora de ganancia/p" & ChrW(233) & "rdida por enajenaci" & ChrW(243) & "n de activos.")
    r = r + 1

    ' =============================================
    ' HOJAS DEL LIBRO
    ' =============================================
    r = WriteHeader(ws, r, "Hojas del libro")
    r = WriteText(ws, r, "Este libro contiene las siguientes hojas de trabajo:")
    r = r + 1

    r = WriteBullet(ws, r, "Instrucciones", "Esta hoja. Gu" & ChrW(237) & "a de uso paso a paso.")
    r = WriteBullet(ws, r, "Cat" & ChrW(225) & "logo", "Porcentajes de deducci" & ChrW(243) & "n por tipo de bien (Art. 33, 34 y 35 LISR).")
    r = WriteBullet(ws, r, "Inversiones", "Hoja principal. Registra activos y calcula la deducci" & ChrW(243) & "n actualizada.")
    r = WriteBullet(ws, r, "Resumen", "Totales agrupados por tipo de bien (SUMIF autom" & ChrW(225) & "tico).")
    r = WriteBullet(ws, r, "Baja_Activos", "Calculadora de ganancia o p" & ChrW(233) & "rdida por venta de activos.")
    r = WriteBullet(ws, r, "INPC", ChrW(205) & "ndices Nacionales de Precios al Consumidor (1984-2025).")
    r = WriteBullet(ws, r, "Config", "Par" & ChrW(225) & "metros: ejercicio fiscal y topes de deducibilidad.")
    r = r + 1

    ' =============================================
    ' FLUJO DE TRABAJO
    ' =============================================
    r = WriteHeader(ws, r, "Flujo de trabajo")
    r = r + 1

    ' Paso 1
    r = WriteStep(ws, r, "1", "Configurar el ejercicio fiscal")
    r = WriteText(ws, r, "Ve a la hoja ""Config"" y verifica que el campo ""Ejercicio"" tenga el a" & ChrW(241) & "o correcto.")
    r = WriteText(ws, r, "Este valor se usa en todas las f" & ChrW(243) & "rmulas de la hoja Inversiones.")
    r = r + 1

    ' Paso 2
    r = WriteStep(ws, r, "2", "Registrar tus inversiones")
    r = WriteText(ws, r, "En la hoja ""Inversiones"", llena las columnas de captura manual:")
    r = r + 1
    r = WriteBullet(ws, r, "No.", "N" & ChrW(250) & "mero de control o cuenta contable.")
    r = WriteBullet(ws, r, "Concepto", "Descripci" & ChrW(243) & "n del bien (ej: ""Laptop Dell Latitude 5540"").")
    r = WriteBullet(ws, r, "Fecha de Adquisici" & ChrW(243) & "n", "Fecha en que se adquiri" & ChrW(243) & " el bien.")
    r = WriteBullet(ws, r, "M.O.I.", "Monto Original de la Inversi" & ChrW(243) & "n (precio + fletes + instalaci" & ChrW(243) & "n, sin IVA).")
    r = WriteBullet(ws, r, "Tipo de Bien", "Selecciona del men" & ChrW(250) & " desplegable (cat" & ChrW(225) & "logo Art. 33/34/35).")
    r = WriteBullet(ws, r, "Dep. Acumulada", "Columna O. Captura manual. Ver nota abajo.")
    r = r + 1
    r = WriteNote(ws, r, "Las dem" & ChrW(225) & "s columnas se calculan autom" & ChrW(225) & "ticamente con f" & ChrW(243) & "rmulas.")
    r = r + 1

    ' Nota especial sobre Dep. Acumulada
    r = WriteStep(ws, r, "2b", "Sobre la columna Dep. Acumulada (col. O)")
    r = WriteText(ws, r, "Esta columna es de CAPTURA MANUAL. Representa la suma de todas las deducciones")
    r = WriteText(ws, r, "fiscales que ya aplicaste en ejercicios anteriores para cada activo.")
    r = r + 1
    r = WriteText(ws, r, "De d" & ChrW(243) & "nde obtener este dato:")
    r = WriteBullet(ws, r, "Balanza de comprobaci" & ChrW(243) & "n", "Saldo acumulado de la cuenta de depreciaci" & ChrW(243) & "n fiscal acumulada.")
    r = WriteBullet(ws, r, "Papeles de trabajo", "Si llevas el control en Excel, es la suma de col. J de a" & ChrW(241) & "os anteriores.")
    r = WriteBullet(ws, r, "Declaraci" & ChrW(243) & "n anual anterior", "El monto de deducci" & ChrW(243) & "n de inversiones declarado en ejercicios previos.")
    r = WriteBullet(ws, r, "Esta misma plantilla", "Si usaste esta plantilla el a" & ChrW(241) & "o pasado, toma Dep. Acum + Deducci" & ChrW(243) & "n del Ejercicio.")
    r = r + 1
    r = WriteNote(ws, r, "Para activos NUEVOS (adquiridos en el ejercicio actual), deja esta columna en cero o vac" & ChrW(237) & "a.")
    r = WriteNote(ws, r, "Para activos que YA SE DEPRECIARON COMPLETAMENTE, la Dep. Acumulada debe ser igual al MOI Deducible.")
    r = WriteNote(ws, r, "Si este dato es incorrecto, el Saldo Pendiente de Deducir (col. P) no ser" & ChrW(225) & " confiable.")
    r = r + 1

    ' Paso 3
    r = WriteStep(ws, r, "3", "Revisar c" & ChrW(225) & "lculos autom" & ChrW(225) & "ticos")
    r = WriteText(ws, r, "Las siguientes columnas se calculan solas al llenar los datos:")
    r = r + 1
    r = WriteBullet(ws, r, "MOI Deducible", "Aplica topes de autom" & ChrW(243) & "viles autom" & ChrW(225) & "ticamente ($175K/$250K).")
    r = WriteBullet(ws, r, "% Deducci" & ChrW(243) & "n", "Se obtiene del cat" & ChrW(225) & "logo seg" & ChrW(250) & "n el tipo de bien seleccionado.")
    r = WriteBullet(ws, r, "Meses de Uso", "Meses completos de uso en el ejercicio.")
    r = WriteBullet(ws, r, "Deducci" & ChrW(243) & "n del Ejercicio", "= MOI Deducible x % x Meses/12.")
    r = WriteBullet(ws, r, "INPC / Factor", "Actualizaci" & ChrW(243) & "n por inflaci" & ChrW(243) & "n (4 decimales, Art. 9 RLISR).")
    r = WriteBullet(ws, r, "Deducci" & ChrW(243) & "n Actualizada", "= Deducci" & ChrW(243) & "n x Factor de Actualizaci" & ChrW(243) & "n.")
    r = WriteBullet(ws, r, "Saldo Pendiente", "= MOI Deducible - Dep. Acumulada - Deducci" & ChrW(243) & "n del Ejercicio.")
    r = r + 1

    ' Paso 4
    r = WriteStep(ws, r, "4", "Consultar el resumen")
    r = WriteText(ws, r, "La hoja ""Resumen"" muestra los totales por tipo de bien autom" & ChrW(225) & "ticamente.")
    r = WriteText(ws, r, ChrW(218) & "til para declaraciones anuales y reportes financieros.")
    r = r + 1

    ' Paso 5
    r = WriteStep(ws, r, "5", "Calcular bajas de activos (si aplica)")
    r = WriteText(ws, r, "Si vendes o das de baja un activo, usa la hoja ""Baja_Activos"".")
    r = WriteText(ws, r, "Ingresa el MOI, deducciones acumuladas, INPCs y precio de venta.")
    r = WriteText(ws, r, "La hoja calcula si hay ganancia acumulable o p" & ChrW(233) & "rdida deducible.")
    r = r + 1

    ' =============================================
    ' TOPES DE AUTOMOVILES
    ' =============================================
    r = WriteHeader(ws, r, "Topes de autom" & ChrW(243) & "viles (Art. 36 LISR)")
    r = r + 1

    Dim tblTop As Long
    tblTop = r

    ws.Cells(r, 2).Value = "Tipo de veh" & ChrW(237) & "culo    " & ChrW(8594) & "    Tope deducible"
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 2).Font.Color = vbWhite
    ws.Cells(r, 2).Interior.Color = RGB(44, 62, 80)
    r = r + 1

    r = WriteTableRow(ws, r, "Combusti" & ChrW(243) & "n interna", "$175,000 MXN")
    r = WriteTableRow(ws, r, "El" & ChrW(233) & "ctrico o h" & ChrW(237) & "brido", "$250,000 MXN")
    r = WriteTableRow(ws, r, "Pick-up (cami" & ChrW(243) & "n de carga)", "Sin tope (100% deducible)")

    With ws.Range(ws.Cells(tblTop, 2), ws.Cells(r - 1, 2)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(189, 195, 199)
    End With
    r = r + 1

    r = WriteNote(ws, r, "Los topes se configuran en la hoja Config y se aplican autom" & ChrW(225) & "ticamente en la columna MOI Deducible.")
    r = WriteNote(ws, r, "Las pick-up se clasifican como camiones de carga (Criterio 27/ISR/N) y no tienen tope.")
    r = r + 1

    ' =============================================
    ' NOTAS IMPORTANTES
    ' =============================================
    r = WriteHeader(ws, r, "Notas importantes")
    r = r + 1
    r = WriteNote(ws, r, "Habilitar macros: Al abrir el archivo, haz clic en ""Habilitar contenido"" si aparece la barra de seguridad.")
    r = WriteNote(ws, r, "El IVA NO forma parte del MOI (es acreditable), salvo que no tengas derecho al acreditamiento.")
    r = WriteNote(ws, r, "Si no deduces en el ejercicio de inicio de uso ni en el siguiente, pierdes esos montos de forma permanente.")
    r = WriteNote(ws, r, "Puedes aplicar un porcentaje menor al m" & ChrW(225) & "ximo, pero queda fijo por 5 a" & ChrW(241) & "os (Art. 66 RLISR).")
    r = WriteNote(ws, r, "El Factor de Actualizaci" & ChrW(243) & "n se calcula a 4 decimales conforme al Art. 9 del Reglamento de la LISR.")
    r = WriteNote(ws, r, "La tabla INPC se puede actualizar manualmente agregando filas para a" & ChrW(241) & "os futuros.")
    r = WriteNote(ws, r, "Para bienes de energ" & ChrW(237) & "a renovable (100%), el sistema debe operar al menos 5 a" & ChrW(241) & "os continuos.")
    r = r + 2

    ' =============================================
    ' PIE
    ' =============================================
    With ws.Cells(r, 2)
        .Value = "Versi" & ChrW(243) & "n 1.0 | Marzo 2026"
        .Font.Size = 9
        .Font.Color = RGB(149, 165, 166)
        .Font.Italic = True
    End With

    ' Proteger hoja
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
    Dim pre As String
    Dim fullText As String
    pre = ChrW(8226) & "  "
    fullText = pre & etiqueta & " " & ChrW(8212) & " " & descripcion

    With ws.Cells(r, 2)
        .Value = fullText
        .Font.Size = 11
        .Font.Color = RGB(100, 100, 100)
        .Font.Bold = False
        ' Negrita solo para la etiqueta
        .Characters(Start:=Len(pre) + 1, Length:=Len(etiqueta)).Font.Bold = True
        .Characters(Start:=Len(pre) + 1, Length:=Len(etiqueta)).Font.Color = RGB(44, 62, 80)
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
    ws.Cells(r, 2).Value = col1 & "    " & ChrW(8594) & "    " & col2

    If r Mod 2 = 0 Then
        ws.Cells(r, 2).Interior.Color = RGB(245, 247, 249)
    End If

    WriteTableRow = r + 1
End Function
