Attribute VB_Name = "ModuloExportadorDIOT"
Sub ExportarDIOT()
    Dim ws As Worksheet
    Dim rutaArchivo As String
    Dim stream As Object
    Dim ultimaFila As Long
    Dim ultimaColumna As Long
    Dim fila As Long, col As Long
    Dim linea As String
    Dim nombreHoja As String
    Dim datosFila() As String
    Dim valorCelda As String
    Dim colPais As Long
    Dim rngDatos As Variant

    ' Definir la hoja activa
    Set ws = ActiveSheet
    nombreHoja = ws.Name

    ' Limpiar nombre de hoja para archivo
    Dim chars As Variant, c As Variant
    chars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each c In chars
        nombreHoja = Replace(nombreHoja, c, "_")
    Next c

    rutaArchivo = ThisWorkbook.Path & "\DIOT_" & nombreHoja & "_CargaMasiva.txt"

    ' Verificar si el archivo está abierto
    If Dir(rutaArchivo) <> "" Then
        On Error Resume Next
        Open rutaArchivo For Append Access Write As #1
        If Err.Number <> 0 Then
            MsgBox "El archivo está en uso. Ciérralo e inténtalo nuevamente.", vbExclamation, "Error de Escritura"
            Exit Sub
        End If
        Close #1
        On Error GoTo 0
    End If

    ' Determinar dimensiones
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ultimaColumna = ws.Cells(5, ws.Columns.Count).End(xlToLeft).Column
    
    If ultimaFila < 6 Then
        MsgBox "No hay datos para exportar.", vbExclamation
        Exit Sub
    End If

    ' Identificar columna de país una sola vez
    colPais = 0
    For col = 1 To ultimaColumna
        If UCase(Trim(ws.Cells(5, col).Value)) = "PAÍS O JURISDICCIÓN DE RESIDENCIA FISCAL" Then
            colPais = col
            Exit For
        End If
    Next col

    ' Leer todo el rango a memoria para velocidad masiva
    rngDatos = ws.Range(ws.Cells(6, 1), ws.Cells(ultimaFila, ultimaColumna)).Value

    ' Configurar Stream UTF-8
    On Error Resume Next
    Set stream = CreateObject("ADODB.Stream")
    If Err.Number <> 0 Then
        MsgBox "Error al iniciar el proceso de escritura.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText ChrW(&HFEFF) ' BOM

    ' Procesar datos desde el array
    For fila = 1 To UBound(rngDatos, 1)
        ReDim datosFila(1 To ultimaColumna)
        For col = 1 To ultimaColumna
            valorCelda = Trim(CStr(rngDatos(fila, col)))
            
            ' Procesar código de país si es la columna correcta
            If col = colPais And valorCelda <> "" Then
                datosFila(col) = ObtenerCodigoPais(valorCelda)
            Else
                datosFila(col) = valorCelda
            End If
        Next col
        
        linea = Join(datosFila, "|")
        stream.WriteText linea & vbCrLf
    Next fila

    ' Guardar y Cerrar
    On Error Resume Next
    stream.SaveToFile rutaArchivo, 2
    If Err.Number <> 0 Then
        MsgBox "Error al escribir el archivo.", vbCritical
        Exit Sub
    End If
    stream.Close
    On Error GoTo 0

    MsgBox "Archivo exportado correctamente en: " & rutaArchivo, vbInformation, "Exportación Completa"
End Sub

' Función optimizada para obtener el código ISO ALPHA-3 del país
Function ObtenerCodigoPais(nombrePais As String) As String
    Static dic As Object
    
    ' Inicializar el diccionario solo la primera vez para máximo rendimiento
    If dic Is Nothing Then
        Set dic = CreateObject("Scripting.Dictionary")
        With dic
            .Add "AFGANISTÁN", "AFG"
            .Add "ISLAS ALAND", "ALA"
            .Add "ALBANIA", "ALB"
            .Add "ALEMANIA", "DEU"
            .Add "ANDORRA", "AND"
            .Add "ANGOLA", "AGO"
            .Add "ANGUILA", "AIA"
            .Add "ANTÁRTIDA", "ATA"
            .Add "ANTIGUA Y BARBUDA", "ATG"
            .Add "ARABIA SAUDITA", "SAU"
            .Add "ARGELIA", "DZA"
            .Add "ARGENTINA", "ARG"
            .Add "ARMENIA", "ARM"
            .Add "ARUBA", "ABW"
            .Add "AUSTRALIA", "AUS"
            .Add "AUSTRIA", "AUT"
            .Add "AZERBAIYÁN", "AZE"
            .Add "BAHAMAS (LAS)", "BHS"
            .Add "BANGLADÉS", "BGD"
            .Add "BARBADOS", "BRB"
            .Add "BARÉIN", "BHR"
            .Add "BÉLGICA", "BEL"
            .Add "BELICE", "BLZ"
            .Add "BENÍN", "BEN"
            .Add "BERMUDAS", "BMU"
            .Add "BIELORRUSIA", "BLR"
            .Add "MYANMAR", "MMR"
            .Add "BOLIVIA, ESTADO PLURINACIONAL DE", "BOL"
            .Add "BOSNIA Y HERZEGOVINA", "BIH"
            .Add "BOTSUANA", "BWA"
            .Add "BRASIL", "BRA"
            .Add "BRUNÉI DARUSSALAM", "BRN"
            .Add "BULGARIA", "BGR"
            .Add "BURKINA FASO", "BFA"
            .Add "BURUNDI", "BDI"
            .Add "BUTÁN", "BTN"
            .Add "CABO VERDE", "CPV"
            .Add "CAMBOYA", "KHM"
            .Add "CAMERÚN", "CMR"
            .Add "CANADÁ", "CAN"
            .Add "CATAR", "QAT"
            .Add "BONAIRE, SAN EUSTAQUIO Y SABA", "BES"
            .Add "CHAD", "TCD"
            .Add "CHILE", "CHL"
            .Add "CHINA", "CHN"
            .Add "CHIPRE", "CYP"
            .Add "COLOMBIA", "COL"
            .Add "COMORAS", "COM"
            .Add "COREA (LA REPÚBLICA DEMOCRÁTICA POPULAR DE)", "PRK"
            .Add "COREA (LA REPÚBLICA DE)", "KOR"
            .Add "CÔTE D'IVOIRE", "CIV"
            .Add "COSTA RICA", "CRI"
            .Add "CROACIA", "HRV"
            .Add "CUBA", "CUB"
            .Add "CURAÇAO", "CUW"
            .Add "DINAMARCA", "DNK"
            .Add "DOMINICA", "DMA"
            .Add "ECUADOR", "ECU"
            .Add "EGIPTO", "EGY"
            .Add "EL SALVADOR", "SLV"
            .Add "EMIRATOS ÁRABES UNIDOS (LOS)", "ARE"
            .Add "ERITREA", "ERI"
            .Add "ESLOVAQUIA", "SVK"
            .Add "ESLOVENIA", "SVN"
            .Add "ESPAÑA", "ESP"
            .Add "ESTADOS UNIDOS (LOS)", "USA"
            .Add "ESTONIA", "EST"
            .Add "ETIOPÍA", "ETH"
            .Add "FILIPINAS (LAS)", "PHL"
            .Add "FINLANDIA", "FIN"
            .Add "FIYI", "FJI"
            .Add "FRANCIA", "FRA"
            .Add "GABÓN", "GAB"
            .Add "GAMBIA (LA)", "GMB"
            .Add "GEORGIA", "GEO"
            .Add "GHANA", "GHA"
            .Add "GIBRALTAR", "GIB"
            .Add "GRANADA", "GRD"
            .Add "GRECIA", "GRC"
            .Add "GROENLANDIA", "GRL"
            .Add "GUADALUPE", "GLP"
            .Add "GUAM", "GUM"
            .Add "GUATEMALA", "GTM"
            .Add "GUAYANA FRANCESA", "GUF"
            .Add "GUERNSEY", "GGY"
            .Add "GUINEA", "GIN"
            .Add "GUINEA-BISÁU", "GNB"
            .Add "GUINEA ECUATORIAL", "GNQ"
            .Add "GUYANA", "GUY"
            .Add "HAITÍ", "HTI"
            .Add "HONDURAS", "HND"
            .Add "HONG KONG", "HKG"
            .Add "HUNGRÍA", "HUN"
            .Add "INDIA", "IND"
            .Add "INDONESIA", "IDN"
            .Add "IRAK", "IRQ"
            .Add "IRÁN (LA REPÚBLICA ISLÁMICA DE)", "IRN"
            .Add "IRLANDA", "IRL"
            .Add "ISLA BOUVET", "BVT"
            .Add "ISLA DE MAN", "IMN"
            .Add "ISLA DE NAVIDAD", "CXR"
            .Add "ISLA NORFOLK", "NFK"
            .Add "ISLANDIA", "ISL"
            .Add "ISLAS CAIMÁN (LAS)", "CYM"
            .Add "ISLAS COCOS (KEELING)", "CCK"
            .Add "ISLAS COOK (LAS)", "COK"
            .Add "ISLAS FEROE (LAS)", "FRO"
            .Add "GEORGIA DEL SUR Y LAS ISLAS SANDWICH DEL SUR", "SGS"
            .Add "ISLA HEARD E ISLAS MCDONALD", "HMD"
            .Add "ISLAS MALVINAS [FALKLAND] (LAS)", "FLK"
            .Add "ISLAS MARIANAS DEL NORTE (LAS)", "MNP"
            .Add "ISLAS MARSHALL (LAS)", "MHL"
            .Add "PITCAIRN", "PCN"
            .Add "ISLAS SALOMÓN (LAS)", "SLB"
            .Add "ISLAS TURCAS Y CAICOS (LAS)", "TCA"
            .Add "ISLAS DE ULTRAMAR MENORES DE ESTADOS UNIDOS (LAS)", "UMI"
            .Add "ISLAS VÍRGENES (BRITÁNICAS)", "VGB"
            .Add "ISLAS VÍRGENES (EE.UU.)", "VIR"
            .Add "ISRAEL", "ISR"
            .Add "ITALIA", "ITA"
            .Add "JAMAICA", "JAM"
            .Add "JAPÓN", "JPN"
            .Add "JERSEY", "JEY"
            .Add "JORDANIA", "JOR"
            .Add "KAZAJISTÁN", "KAZ"
            .Add "KENIA", "KEN"
            .Add "KIRGUISTÁN", "KGZ"
            .Add "KIRIBATI", "KIR"
            .Add "KUWAIT", "KWT"
            .Add "LAO, (LA) REPÚBLICA DEMOCRÁTICA POPULAR", "LAO"
            .Add "LESOTO", "LSO"
            .Add "LETONIA", "LVA"
            .Add "LÍBANO", "LBN"
            .Add "LIBERIA", "LBR"
            .Add "LIBIA", "LBY"
            .Add "LIECHTENSTEIN", "LIE"
            .Add "LITUANIA", "LTU"
            .Add "LUXEMBURGO", "LUX"
            .Add "MACAO", "MAC"
            .Add "MADAGASCAR", "MDG"
            .Add "MALASIA", "MYS"
            .Add "MALAUI", "MWI"
            .Add "MALDIVAS", "MDV"
            .Add "MALÍ", "MLI"
            .Add "MALTA", "MLT"
            .Add "MARRUECOS", "MAR"
            .Add "MARTINICA", "MTQ"
            .Add "MAURICIO", "MUS"
            .Add "MAURITANIA", "MRT"
            .Add "MAYOTTE", "MYT"
            .Add "MICRONESIA (LOS ESTADOS FEDERADOS DE)", "FSM"
            .Add "MOLDAVIA (LA REPÚBLICA DE)", "MDA"
            .Add "MÓNACO", "MCO"
            .Add "MONGOLIA", "MNG"
            .Add "MONTENEGRO", "MNE"
            .Add "MONTSERRAT", "MSR"
            .Add "MOZAMBIQUE", "MOZ"
            .Add "NAMIBIA", "NAM"
            .Add "NAURU", "NRU"
            .Add "NEPAL", "NPL"
            .Add "NICARAGUA", "NIC"
            .Add "NÍGER (EL)", "NER"
            .Add "NIGERIA", "NGA"
            .Add "NIUE", "NIU"
            .Add "NORUEGA", "NOR"
            .Add "NUEVA CALEDONIA", "NCL"
            .Add "NUEVA ZELANDA", "NZL"
            .Add "OMÁN", "OMN"
            .Add "PAÍSES BAJOS (LOS)", "NLD"
            .Add "PAKISTÁN", "PAK"
            .Add "PALAOS", "PLW"
            .Add "PALESTINA, ESTADO DE", "PSE"
            .Add "PANAMÁ", "PAN"
            .Add "PAPÚA NUEVA GUINEA", "PNG"
            .Add "PARAGUAY", "PRY"
            .Add "PERÚ", "PER"
            .Add "POLINESIA FRANCESA", "PYF"
            .Add "POLONIA", "POL"
            .Add "PORTUGAL", "PRT"
            .Add "PUERTO RICO", "PRI"
            .Add "REINO UNIDO (EL)", "GBR"
            .Add "REPÚBLICA CENTROAFRICANA (LA)", "CAF"
            .Add "REPÚBLICA CHECA (LA)", "CZE"
            .Add "MACEDONIA (LA ANTIGUA REPÚBLICA YUGOSLAVA DE)", "MKD"
            .Add "CONGO", "COG"
            .Add "CONGO (LA REPÚBLICA DEMOCRÁTICA DEL)", "COD"
            .Add "REPÚBLICA DOMINICANA (LA)", "DOM"
            .Add "REUNIÓN", "REU"
            .Add "RUANDA", "RWA"
            .Add "RUMANIA", "ROU"
            .Add "RUSIA, (LA) FEDERACIÓN DE", "RUS"
            .Add "SAHARA OCCIDENTAL", "ESH"
            .Add "SAMOA", "WSM"
            .Add "SAMOA AMERICANA", "ASM"
            .Add "SAN BARTOLOMÉ", "BLM"
            .Add "SAN CRISTÓBAL Y NIEVES", "KNA"
            .Add "SAN MARINO", "SMR"
            .Add "SAN MARTÍN (PARTE FRANCESA)", "MAF"
            .Add "SAN PEDRO Y MIQUELÓN", "SPM"
            .Add "SAN VICENTE Y LAS GRANADINAS", "VCT"
            .Add "SANTA HELENA, ASCENSIÓN Y TRISTÁN DEACUÑA", "SHN"
            .Add "SANTA LUCÍA", "LCA"
            .Add "SANTO TOMÉ Y PRÍNCIPE", "STP"
            .Add "SENEGAL", "SEN"
            .Add "SERBIA", "SRB"
            .Add "SEYCHELLES", "SYC"
            .Add "SIERRA LEONA", "SLE"
            .Add "SINGAPUR", "SGP"
            .Add "SINT MAARTEN (PARTE HOLANDESA)", "SXM"
            .Add "SIRIA, (LA) REPÚBLICA ÁRABE", "SYR"
            .Add "SOMALIA", "SOM"
            .Add "SRI LANKA", "LKA"
            .Add "SUAZILANDIA", "SWZ"
            .Add "SUDÁFRICA", "ZAF"
            .Add "SUDÁN (EL)", "SDN"
            .Add "SUDÁN DEL SUR", "SSD"
            .Add "SUECIA", "SWE"
            .Add "SUIZA", "CHE"
            .Add "SURINAM", "SUR"
            .Add "SVALBARD Y JAN MAYEN", "SJM"
            .Add "TAILANDIA", "THA"
            .Add "TAIWÁN (PROVINCIA DE CHINA)", "TWN"
            .Add "TANZANIA, REPÚBLICA UNIDA DE", "TZA"
            .Add "TAYIKISTÁN", "TJK"
            .Add "TERRITORIO BRITÁNICO DEL OCÉANO ÍNDICO (EL)", "IOT"
            .Add "TERRITORIOS AUSTRALES FRANCESES (LOS)", "ATF"
            .Add "TIMOR-LESTE", "TLS"
            .Add "TOGO", "TGO"
            .Add "TOKELAU", "TKL"
            .Add "TONGA", "TON"
            .Add "TRINIDAD Y TOBAGO", "TTO"
            .Add "TÚNEZ", "TUN"
            .Add "TURKMENISTÁN", "TKM"
            .Add "TURQUÍA", "TUR"
            .Add "TUVALU", "TUV"
            .Add "UCRANIA", "UKR"
            .Add "UGANDA", "UGA"
            .Add "URUGUAY", "URY"
            .Add "UZBEKISTÁN", "UZB"
            .Add "VANUATU", "VUT"
            .Add "SANTA SEDE[ESTADO VATICANO] (LA) DE LA CIUDAD DEL", "VAT"
            .Add "VENEZUELA, REPÚBLICA BOLIVARIANA DE", "VEN"
            .Add "VIET NAM", "VNM"
            .Add "WALLIS Y FUTUNA", "WLF"
            .Add "YEMEN", "YEM"
            .Add "YIBUTI", "DJI"
            .Add "ZAMBIA", "ZMB"
            .Add "ZIMBABUE", "ZWE"
            .Add "OTRO", "ZZZ"
        End With
    End If

    ' Buscar en el diccionario
    Dim busqueda As String
    busqueda = UCase(Trim(nombrePais))
    
    If dic.exists(busqueda) Then
        ObtenerCodigoPais = dic(busqueda)
    Else
        ObtenerCodigoPais = nombrePais
    End If
End Function



