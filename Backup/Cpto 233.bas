Attribute VB_Name = "Módulo2"
Sub Filtrar_Docuementos()
    Dim contenido As String
    Dim wbContenido As Workbook, _
        wsResultado As Excel.Worksheet, _
        wsError As Excel.Worksheet, _
        wsContenido As Excel.Worksheet
    Dim nFilas As Double
    Dim nColumnas As Integer
    Dim rango As Range
    Dim nFilasCont As Double
    Dim rangoCont As Range
    Dim rangoError As Range
    Dim rangoTemp As String
    Dim i As Integer
    Dim j As Long
    Dim m As Integer
    Dim valorJur As Integer
    Dim valorDoc As String
    Dim valorAct As String
    Dim resultado As Range
    Dim celdaDoc As String
    Dim temp As String
    Dim nFilasResultado As Integer
    Dim nFilasError As Integer
    Dim nColumnasError As Integer
    Dim bandera As Boolean
    Dim totalImporte As Double

    
    'Indicar el libro de excel CONTENIDO y control de errores
    contenido = InputBox("Ingrese el nombre del archivo:", "Abrir", "Archivo.xlsx")
    If contenido <> "" Then
        On Error GoTo ControlErrorOpen
        Set wbContenido = Workbooks.Open(ActiveWorkbook.Path & "\" & contenido)
    Else
        Exit Sub
    End If
    
    'Activar este libro
    ThisWorkbook.Activate
    
    'Borra las hojas destino si existen
    Application.DisplayAlerts = False
    
    'Agrega las nuevas hojas
    Worksheets.Add
    ActiveSheet.Name = "Resultados"
    Worksheets.Add
    ActiveSheet.Name = "Errores"
    Application.DisplayAlerts = True
    
    Set wsResultado = Worksheets("Resultados")
    Set wsError = Worksheets("Errores")
    Set wsContenido = wbContenido.Worksheets("Detalle x Agente")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    'Fila del encabezado Error
    Range("1:1").Copy
    wsError.Range("1:1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    nFilasError = 2
    Set rangoError = wsError.UsedRange
    nColumnasError = rangoError.Columns.Count
    nColumnasError = nColumnasError + 1
    wsError.Range("A1").Copy
    wsError.Cells(1, nColumnasError).PasteSpecial Paste:=xlFormats
    Application.CutCopyMode = False
    wsError.Cells(1, nColumnasError).Value = "Mensaje"
    wsError.Columns(nColumnasError).ColumnWidth = 52
    
    'Fila del encabezado Resultado
    wsContenido.Range("A1").Copy
    wsResultado.Range("A1:L1").PasteSpecial Paste:=xlFormats
    Application.CutCopyMode = False
    nFilasResultado = 1
    wsResultado.Cells(nFilasResultado, 1).Value = "PtaId"
    wsResultado.Cells(nFilasResultado, 2).Value = "JurId"
    wsResultado.Cells(nFilasResultado, 3).Value = "EscId"
    wsResultado.Cells(nFilasResultado, 4).Value = "Pref"
    wsResultado.Cells(nFilasResultado, 5).Value = "Doc"
    wsResultado.Cells(nFilasResultado, 6).Value = "Digito"
    wsResultado.Cells(nFilasResultado, 7).Value = "Nombres"
    wsResultado.Cells(nFilasResultado, 8).Value = "Couc"
    wsResultado.Cells(nFilasResultado, 9).Value = "Reajuste"
    wsResultado.Cells(nFilasResultado, 10).Value = "Unidades"
    wsResultado.Cells(nFilasResultado, 11).Value = "Importe"
    wsResultado.Cells(nFilasResultado, 12).Value = "Vto"
    wsResultado.Range("A1").Copy
    wsResultado.Cells(nFilasResultado, 13).PasteSpecial Paste:=xlFormats
    wsResultado.Cells(nFilasResultado, 14).PasteSpecial Paste:=xlFormats
    Application.CutCopyMode = False
    wsResultado.Cells(nFilasResultado, 13).Value = "Totales"
    wsResultado.Cells(nFilasResultado, 14).Value = "Año"
    wsResultado.Cells(nFilasResultado, 15).Value = "Actuación"
    wsResultado.Cells(nFilasResultado, 16).Value = "Proceso"
    nFilasResultado = 2
    
    Cells(1, nColumnas + 1).Value = "Observación"
    
    For i = 2 To nFilas
        If Cells(i, nColumnas).Value = "" Then
            valorJur = Cells(i, 1).Value
            valorDoc = Cells(i, 3).Value
            valorAct = Cells(i, 5).Value
            totalImporte = 0
            'Busca en el otro archivo
            rangoTemp = "D2:D" & nFilasCont
            Set resultado = wsContenido.Range(rangoTemp).Find(What:=valorDoc, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False, _
                        SearchFormat:=False)
            'Si el resultado de la búsqueda no es vacío
            If Not resultado Is Nothing Then
                primerResultado = resultado.Address
                'Se obtiene el valor de j
                celdaDoc = resultado.Address
                tempDoc = ""
                For m = 1 To Len(celdaDoc)
                    If IsNumeric(Mid(celdaDoc, m, 1)) Then
                        tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                    End If
                Next m
                j = tempDoc
                If wsContenido.Cells(j, 1).Value = valorJur Then
                    'retrocedo hasta encontrar el primero
                    bandera = True
                    Do While bandera
                        If wsContenido.Cells(j - 1, 4).Value = valorDoc Then
                            j = j - 1
                        Else
                            bandera = False
                        End If
                    Loop
                    
                    'controlar nombre
                    If Mid(wsContenido.Cells(j, 6).Value, 1, 3) <> Mid(Cells(i, 4).Value, 1, 3) Then
                        Cells(i, nColumnas + 2).Value = "Controlar DNI+Nombre"
                    End If
                    
                    Do
                        If wsContenido.Cells(j, 19).Value > 0 Then
                            'Copio y pego la fila de Contenido a Resultado
                            wsResultado.Cells(nFilasResultado, 1).Value = wsContenido.Cells(j, 2).Value
                            wsResultado.Cells(nFilasResultado, 2).Value = wsContenido.Cells(j, 1).Value
                            wsResultado.Cells(nFilasResultado, 3).Value = wsContenido.Cells(j, 7).Value
                            wsResultado.Cells(nFilasResultado, 4).Value = wsContenido.Cells(j, 3).Value
                            wsResultado.Cells(nFilasResultado, 5).Value = wsContenido.Cells(j, 4).Value
                            wsResultado.Cells(nFilasResultado, 6).Value = wsContenido.Cells(j, 5).Value
                            wsResultado.Cells(nFilasResultado, 7).Value = wsContenido.Cells(j, 6).Value
                            wsResultado.Cells(nFilasResultado, 8).Value = wsContenido.Cells(j, 15).Value
                            wsResultado.Cells(nFilasResultado, 9).Value = 1
                            wsResultado.Cells(nFilasResultado, 10).Value = 0
                            wsResultado.Cells(nFilasResultado, 11).Value = wsContenido.Cells(j, 19).Value
                            wsResultado.Cells(nFilasResultado, 12).Value = wsContenido.Cells(j, 20).Value
                            totalImporte = totalImporte + wsContenido.Cells(j, 19).Value
                            wsResultado.Cells(nFilasResultado, 14).Value = Cells(i, 5).Value
                            wsResultado.Cells(nFilasResultado, 15).Value = Cells(i, 6).Value
                            wsResultado.Cells(nFilasResultado, 16).Value = Cells(i, 7).Value
                            'Aumento el valor del contador de las filas
                            nFilasResultado = nFilasResultado + 1
                        End If
                        j = j + 1
                    Loop While valorDoc = wsContenido.Cells(j, 4).Value
                    If totalImporte <> 0 Then
                        wsResultado.Cells(nFilasResultado - 1, 13).Value = totalImporte
                        wsResultado.Cells(nFilasResultado - 1, 13).Font.Bold = True
                    Else
                        'Importe total igual a 0
                        'Copio la fila de Origen
                        temp = i & ":" & i
                        Range(temp).Copy
                        'Pegar y actualizar num de filas. Agregando el msj correspondiente
                        temp = nFilasError & ":" & nFilasError
                        wsError.Range(temp).PasteSpecial xlPasteAll
                        Application.CutCopyMode = False
                        wsError.Cells(nFilasError, nColumnasError).Value = "El importe total es 0."
                        nFilasError = nFilasError + 1
                        
                        Cells(i, nColumnas + 1).Value = "Pagado"
                    End If
                Else
                    'No se encontró el documento en la jurisdicción
                    'Copio la fila de Origen
                    temp = i & ":" & i
                    Range(temp).Copy
                    'Pegar y actualizar num de filas. Agregando el msj correspondiente
                    temp = nFilasError & ":" & nFilasError
                    wsError.Range(temp).PasteSpecial xlPasteAll
                    Application.CutCopyMode = False
                    wsError.Cells(nFilasError, nColumnasError).Value = "No se encontró el Documento en la Jurisdicción indicada. Está en la " & wsContenido.Cells(j, 1).Value
                    nFilasError = nFilasError + 1
                    
                    Cells(i, nColumnas + 1).Value = "No se encontró el DNI"
                End If
            Else
                'No se encontró el documento
                'Copio la fila de Origen
                temp = i & ":" & i
                Range(temp).Copy
                'Pegar y actualizar num de filas. Agregando el msj correspondiente
                temp = nFilasError & ":" & nFilasError
                wsError.Range(temp).PasteSpecial xlPasteAll
                Application.CutCopyMode = False
                wsError.Cells(nFilasError, nColumnasError).Value = "No se encontró el Documento."
                nFilasError = nFilasError + 1
                
                Cells(i, nColumnas + 1).Value = "No se encontró el DNI"
            End If
        Else
            temp = i & ":" & i
            Range(temp).Copy
            'Pegar y actualizar num de filas. Agregando el msj correspondiente
            temp = nFilasError & ":" & nFilasError
            wsError.Range(temp).PasteSpecial xlPasteAll
            Application.CutCopyMode = False
            wsError.Cells(nFilasError, nColumnasError).Value = "Pagado"
            nFilasError = nFilasError + 1
            
            Cells(i, nColumnas + 1).Value = Cells(i, nColumnas).Value
        End If
    Next i
    
    'Ver si funciona el cambio de nombre
    ActiveSheet.Name = "Actuaciones"
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


Sub Filtrar_Docuementos_Nuevo()
    Dim contenido As String
    Dim wbContenido As Workbook, _
        wsResultado As Excel.Worksheet, _
        wsContenido As Excel.Worksheet
    Dim nFilas As Double
    Dim nColumnas As Integer
    Dim rango As Range
    Dim nFilasCont As Double
    Dim rangoCont As Range
    Dim i As Integer
    Dim j As Long
    Dim m As Integer
    Dim valorJur As Integer
    Dim valorDoc As String
    Dim valorAct As String
    Dim resultado As Range
    Dim celdaDoc As String
    Dim temp As String
    Dim nFilasResultado As Integer
    Dim totalImporte As Double

    
    'Indicar el libro de excel CONTENIDO y control de errores
    contenido = InputBox("Ingrese el nombre del archivo:", "Abrir", "Archivo.xlsx")
    If contenido <> "" Then
        On Error GoTo ControlErrorOpen
        Set wbContenido = Workbooks.Open(ActiveWorkbook.Path & "\" & contenido)
    Else
        Exit Sub
    End If
    
    'Activar este libro
    ThisWorkbook.Activate
    
    'Borra las hojas destino si existen
    Application.DisplayAlerts = False
    
    'Agrega las nuevas hojas
    Worksheets.Add
    ActiveSheet.Name = "Resultados"
    Application.DisplayAlerts = True
    
    Set wsResultado = Worksheets("Resultados")
    Set wsContenido = wbContenido.Worksheets("Resultado")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    'Fila del encabezado Resultado
    wsContenido.Range("1:1").Copy
    wsResultado.Range("1:1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    nFilasResultado = 2
    
    Cells(1, nColumnas + 1).Value = "Observación"
    
    For i = 2 To nFilas
        valorJur = Cells(i, 1).Value
        valorDoc = Cells(i, 3).Value
        valorAct = Cells(i, 5).Value
        totalImporte = 0
        'Busca en el otro archivo
        rangoTemp = "E2:E" & nFilasCont
        Set resultado = wsContenido.Range(rangoTemp).Find(What:=valorDoc, _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False, _
                    SearchFormat:=False)
        'Si el resultado de la búsqueda no es vacío
        If Not resultado Is Nothing Then
            primerResultado = resultado.Address
            'Se obtiene el valor de j
            celdaDoc = resultado.Address
            tempDoc = ""
            For m = 1 To Len(celdaDoc)
                If IsNumeric(Mid(celdaDoc, m, 1)) Then
                    tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                End If
            Next m
            j = tempDoc
            
            'retrocedo hasta encontrar el primero
            bandera = True
            Do While bandera
                If wsContenido.Cells(j - 1, 4).Value = valorDoc Then
                    j = j - 1
                Else
                    bandera = False
                End If
            Loop
            
            'controlar nombre
            If Mid(wsContenido.Cells(j, 6).Value, 1, 3) <> Mid(Cells(i, 4).Value, 1, 3) Then
                Cells(i, nColumnas + 2).Value = "Controlar DNI+Nombre"
            End If
            
            Do
                wsContenido.Range(j & ":" & j).Copy
                wsResultado.Range(nFilasResultado & ":" & nFilasResultado).PasteSpecial xlPasteAll
                Application.CutCopyMode = False
                
                totalImporte = totalImporte + wsContenido.Cells(j, 11).Value
                wsResultado.Cells(nFilasResultado, 13).Value = ""
                nFilasResultado = nFilasResultado + 1
                j = j + 1
            Loop While valorDoc = wsContenido.Cells(j, 5).Value
            
            wsResultado.Cells(nFilasResultado - 1, 13).Value = totalImporte
            
        Else
            'No se encontró el documento
            Cells(i, nColumnas + 1).Value = "No se encontró el DNI"
        End If
    Next i
    
    'Ver si funciona el cambio de nombre
    ActiveSheet.Name = "Actuaciones"
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


Sub Controlar_Pago()
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet
    Dim valorDoc As String
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim nColumnasCont As Long
    Dim nFilasCont As Long
    Dim i As Long
    Dim contador As Integer
    Dim mes As Integer
    Dim jur As Integer
    Dim anio As Integer
    Dim importeTotal As Double
    
    
    'Indicar el libro de excel CONTENIDO y control de errores
    contenido = InputBox("Ingrese el nombre del archivo:", "Abrir", "Archivo.xlsx")
    If contenido <> "" Then
        On Error GoTo ControlErrorOpen
        Set wbContenido = Workbooks.Open(ActiveWorkbook.Path & "\" & contenido)
    Else
        Exit Sub
    End If
    
    'Activar este libro
    ThisWorkbook.Activate

    Set wsContenido = wbContenido.Worksheets("Hoja1")
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select

    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    nColumnasCont = rangoCont.Columns.Count
    
    Cells(1, nColumnas + 1).Value = "Pagado"
    Cells(1, nColumnas + 2).Value = "Importe"
    
    For i = 2 To nFilas
        valorJur = Cells(i, 1).Value
        valorDoc = Cells(i, 3).Value
        'Busca en el otro archivo
        rangoTemp = "H2:H" & nFilasCont
        Set resultado = wsContenido.Range(rangoTemp).Find(What:=valorDoc, _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False, _
                    SearchFormat:=False)
        'Si el resultado de la búsqueda no es vacío
        If Not resultado Is Nothing Then
            primerResultado = resultado.Address
            'Se obtiene el valor de j
            celdaDoc = resultado.Address
            tempDoc = ""
            For m = 1 To Len(celdaDoc)
                If IsNumeric(Mid(celdaDoc, m, 1)) Then
                    tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                End If
            Next m
            j = tempDoc
            
            Do
                If wsContenido.Cells(j, 7).Value = valorJur Then
                    Cells(i, nColumnas + 1).Value = Cells(i, nColumnas + 1).Value + 1
                    Cells(i, nColumnas + 3).Value = Cells(i, nColumnas + 3).Value + 1
                    If wsContenido.Cells(j, 5).Value = 2 Then
                        Cells(i, nColumnas + 2).Value = Cells(i, nColumnas + 2).Value - wsContenido.Cells(j, 6).Value
                        Cells(i, nColumnas + 3).Value = Cells(i, nColumnas + 3).Value - 1
                    Else
                        Cells(i, nColumnas + 2).Value = wsContenido.Cells(j, 6).Value + Cells(i, nColumnas + 2).Value
                    End If
                End If
                j = j + 1
            Loop While valorDoc = wsContenido.Cells(j, 8).Value
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


Sub Pago_CEIC()
    Dim contenido As String
    Dim wbContenido As Workbook, _
        wsResultado As Excel.Worksheet, _
        wsContenido As Excel.Worksheet
    Dim nFilas As Double
    Dim nColumnas As Integer
    Dim rango As Range
    Dim nFilasCont As Double
    Dim rangoCont As Range
    Dim rangoTemp As String
    Dim i As Integer
    Dim j As Long
    Dim m As Integer
    Dim valorCargo As String
    Dim valorDesde As String
    Dim valorHasta As String
    Dim resultado As Range
    Dim celdaDoc As String
    Dim temp As String
    Dim nFilasResultado As Integer
    Dim totalImporte As Double

    
    'Indicar el libro de excel CONTENIDO y control de errores
    contenido = InputBox("Ingrese el nombre del archivo:", "Abrir", "Archivo.xlsx")
    If contenido <> "" Then
        On Error GoTo ControlErrorOpen
        Set wbContenido = Workbooks.Open(ActiveWorkbook.Path & "\" & contenido)
    Else
        Exit Sub
    End If
    
    'Activar este libro
    ThisWorkbook.Activate
    
    'Borra las hojas destino si existen
    Application.DisplayAlerts = False
    
    'Agrega las nuevas hojas
    Worksheets.Add
    ActiveSheet.Name = "Resultados"
    Application.DisplayAlerts = True
    
    Set wsResultado = Worksheets("Resultados")
    Set wsContenido = wbContenido.Worksheets("Totales")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    'Fila del encabezado Resultado
    nFilasResultado = 1
    wsResultado.Cells(nFilasResultado, 1).Value = "PtaId"
    wsResultado.Cells(nFilasResultado, 2).Value = "JurId"
    wsResultado.Cells(nFilasResultado, 3).Value = "EscId"
    wsResultado.Cells(nFilasResultado, 4).Value = "Pref"
    wsResultado.Cells(nFilasResultado, 5).Value = "Doc"
    wsResultado.Cells(nFilasResultado, 6).Value = "Digito"
    wsResultado.Cells(nFilasResultado, 7).Value = "Nombres"
    wsResultado.Cells(nFilasResultado, 8).Value = "Couc"
    wsResultado.Cells(nFilasResultado, 9).Value = "Reajuste"
    wsResultado.Cells(nFilasResultado, 10).Value = "Unidades"
    wsResultado.Cells(nFilasResultado, 11).Value = "Importe"
    wsResultado.Cells(nFilasResultado, 12).Value = "Vto"
    wsResultado.Cells(nFilasResultado, 13).Value = "Totales"
    wsResultado.Cells(nFilasResultado, 14).Value = "Actuación"
    nFilasResultado = 2
    
    For i = 2 To nFilas
        valorCargo = Cells(i, 5).Value
        valorDesde = Cells(i, 6).Value
        valorHasta = Cells(i, 7).Value
        totalImporte = 0
        'Busca en el otro archivo
        rangoTemp = "A2:A" & nFilasCont
        Set resultado = wsContenido.Range(rangoTemp).Find(What:=valorCargo, _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False, _
                    SearchFormat:=False)
        'Si el resultado de la búsqueda no es vacío
        If Not resultado Is Nothing Then
            primerResultado = resultado.Address
            'Se obtiene el valor de j
            celdaDoc = resultado.Address
            tempDoc = ""
            For m = 1 To Len(celdaDoc)
                If IsNumeric(Mid(celdaDoc, m, 1)) Then
                    tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                End If
            Next m
            j = tempDoc
            
            Do While wsContenido.Cells(j - 1, 1).Value = valorCargo
                j = j - 1
            Loop
            
            Do While wsContenido.Cells(j, 5).Value <> valorDesde
                j = j + 1
            Loop
            
            Do While wsContenido.Cells(j, 5).Value <> valorHasta
                If wsContenido.Cells(j, 1).Value = valorCargo Then
                    wsResultado.Cells(nFilasResultado, 1).Value = 0
                    wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 1).Value
                    wsResultado.Cells(nFilasResultado, 3).Value = Cells(i, 2).Value
                    wsResultado.Cells(nFilasResultado, 4).Value = 0
                    wsResultado.Cells(nFilasResultado, 5).Value = Cells(i, 3).Value
                    wsResultado.Cells(nFilasResultado, 6).Value = 0
                    wsResultado.Cells(nFilasResultado, 7).Value = Cells(i, 4).Value
                    wsResultado.Cells(nFilasResultado, 8).Value = 233
                    wsResultado.Cells(nFilasResultado, 9).Value = 1
                    wsResultado.Cells(nFilasResultado, 10).Value = 0
                    wsResultado.Cells(nFilasResultado, 11).Value = wsContenido.Cells(j, 4).Value
                    wsResultado.Cells(nFilasResultado, 12).Value = wsContenido.Cells(j, 5).Value
                    totalImporte = totalImporte + wsContenido.Cells(j, 4).Value
                    wsResultado.Cells(nFilasResultado, 14).Value = Cells(i, 8).Value
                    'Aumento el valor del contador de las filas
                    nFilasResultado = nFilasResultado + 1
                    
                    'agregar lo de cpto 316
                    If wsContenido.Cells(j, 6).Value = 6 Or wsContenido.Cells(j, 6).Value = 12 Then
                        wsResultado.Cells(nFilasResultado, 1).Value = 0
                        wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 1).Value
                        wsResultado.Cells(nFilasResultado, 3).Value = Cells(i, 2).Value
                        wsResultado.Cells(nFilasResultado, 4).Value = 0
                        wsResultado.Cells(nFilasResultado, 5).Value = Cells(i, 3).Value
                        wsResultado.Cells(nFilasResultado, 6).Value = 0
                        wsResultado.Cells(nFilasResultado, 7).Value = Cells(i, 4).Value
                        wsResultado.Cells(nFilasResultado, 8).Value = 316
                        wsResultado.Cells(nFilasResultado, 9).Value = 1
                        wsResultado.Cells(nFilasResultado, 10).Value = 0
                        wsResultado.Cells(nFilasResultado, 11).Value = wsContenido.Cells(j, 4).Value / 2
                        wsResultado.Cells(nFilasResultado, 12).Value = wsContenido.Cells(j, 5).Value
                        totalImporte = totalImporte + wsContenido.Cells(j, 4).Value / 2
                        wsResultado.Cells(nFilasResultado, 14).Value = Cells(i, 8).Value
                        nFilasResultado = nFilasResultado + 1
                    End If
                End If
                
                j = j + 1
            Loop
            
            'REPITO PORQUE SINO EL ÚLTIMO MES NO LO TOMA
            If wsContenido.Cells(j, 1).Value = valorCargo And wsContenido.Cells(j, 5).Value = valorHasta Then
                wsResultado.Cells(nFilasResultado, 1).Value = 0
                wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 1).Value
                wsResultado.Cells(nFilasResultado, 3).Value = Cells(i, 2).Value
                wsResultado.Cells(nFilasResultado, 4).Value = 0
                wsResultado.Cells(nFilasResultado, 5).Value = Cells(i, 3).Value
                wsResultado.Cells(nFilasResultado, 6).Value = 0
                wsResultado.Cells(nFilasResultado, 7).Value = Cells(i, 4).Value
                wsResultado.Cells(nFilasResultado, 8).Value = 233
                wsResultado.Cells(nFilasResultado, 9).Value = 1
                wsResultado.Cells(nFilasResultado, 10).Value = 0
                wsResultado.Cells(nFilasResultado, 11).Value = wsContenido.Cells(j, 4).Value
                wsResultado.Cells(nFilasResultado, 12).Value = wsContenido.Cells(j, 5).Value
                totalImporte = totalImporte + wsContenido.Cells(j, 4).Value
                wsResultado.Cells(nFilasResultado, 14).Value = Cells(i, 8).Value
                'Aumento el valor del contador de las filas
                nFilasResultado = nFilasResultado + 1
                
                'agregar lo de cpto 316
                If wsContenido.Cells(j, 6).Value = 6 Or wsContenido.Cells(j, 6).Value = 12 Then
                    wsResultado.Cells(nFilasResultado, 1).Value = 0
                    wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 1).Value
                    wsResultado.Cells(nFilasResultado, 3).Value = Cells(i, 2).Value
                    wsResultado.Cells(nFilasResultado, 4).Value = 0
                    wsResultado.Cells(nFilasResultado, 5).Value = Cells(i, 3).Value
                    wsResultado.Cells(nFilasResultado, 6).Value = 0
                    wsResultado.Cells(nFilasResultado, 7).Value = Cells(i, 4).Value
                    wsResultado.Cells(nFilasResultado, 8).Value = 316
                    wsResultado.Cells(nFilasResultado, 9).Value = 1
                    wsResultado.Cells(nFilasResultado, 10).Value = 0
                    wsResultado.Cells(nFilasResultado, 11).Value = wsContenido.Cells(j, 4).Value / 2
                    wsResultado.Cells(nFilasResultado, 12).Value = wsContenido.Cells(j, 5).Value
                    totalImporte = totalImporte + wsContenido.Cells(j, 4).Value / 2
                    wsResultado.Cells(nFilasResultado, 14).Value = Cells(i, 8).Value
                    nFilasResultado = nFilasResultado + 1
                End If
            End If
            
            wsResultado.Cells(nFilasResultado - 1, 13).Value = totalImporte
            wsResultado.Cells(nFilasResultado - 1, 13).Font.Bold = True
            
        Else
            'No se encontró
            Cells(i, nColumnas + 1).Value = "No encontrado"
        End If
    Next i
    
    'ActiveSheet.Name = "Actuaciones"
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


Sub Listar_Agentes_Pagados()
    Dim wsResultado As Excel.Worksheet
    Dim rango As Range
    Dim nFilas As Double
    Dim importeTotal As Double
    Dim cantidad As Integer
    Dim docAnterior As String
    Dim nFilasResultado As Long
    Dim i As Long
    
    
    MsgBox "Debe estar ordenado por DNI.", , "¡Atención!"
    
    Application.DisplayAlerts = False
    'Agrega las nuevas hojas
    Worksheets.Add
    ActiveSheet.Name = "Resultado"
    Application.DisplayAlerts = True
    
    Set wsResultado = Worksheets("Resultado")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    'Fila del encabezado Resultado
    nFilasResultado = 2
    wsResultado.Cells(1, 1).Value = "JUR"
    wsResultado.Cells(1, 2).Value = "DNI"
    wsResultado.Cells(1, 3).Value = "Nombre y Apellido"
    wsResultado.Cells(1, 4).Value = "Cant Liq"
    wsResultado.Cells(1, 5).Value = "Importe Total"
    wsResultado.Cells(1, 6).Value = "Total Liq"
    wsResultado.Range("1:1").Font.Bold = True
    wsResultado.Range("1:1").HorizontalAlignment = xlCenter
    
    'Inicializar
    cantidad = 0
    cant = 0
    importeTotal = 0
    docAnterior = Cells(2, 8).Value
    
    For i = 2 To nFilas
        If docAnterior = Cells(i, 8).Value Then
            cant = cant + 1
            'Acumulo
            If Cells(i, 5).Value = 2 Then
                cantidad = cantidad - 1
                importeTotal = importeTotal - Cells(i, 6).Value
            Else
                cantidad = cantidad + 1
                importeTotal = importeTotal + Cells(i, 6).Value
            End If
        Else
            'Imprimir
            wsResultado.Cells(nFilasResultado, 1).Value = Cells(i - 1, 7).Value
            wsResultado.Cells(nFilasResultado, 2).Value = Cells(i - 1, 8).Value
            wsResultado.Cells(nFilasResultado, 3).Value = Cells(i - 1, 9).Value
            wsResultado.Cells(nFilasResultado, 4).Value = cantidad
            wsResultado.Cells(nFilasResultado, 5).Value = importeTotal
            wsResultado.Cells(nFilasResultado, 6).Value = cant
            nFilasResultado = nFilasResultado + 1
            'Tratar nuevo
            cantidad = 0
            cant = 0
            importeTotal = 0
            docAnterior = Cells(i, 8).Value
            i = i - 1
        End If
    Next i
    'Imprimir lo último
    wsResultado.Cells(nFilasResultado, 1).Value = Cells(i - 1, 7).Value
    wsResultado.Cells(nFilasResultado, 2).Value = Cells(i - 1, 8).Value
    wsResultado.Cells(nFilasResultado, 3).Value = Cells(i - 1, 9).Value
    wsResultado.Cells(nFilasResultado, 4).Value = cantidad
    wsResultado.Cells(nFilasResultado, 5).Value = importeTotal
    wsResultado.Cells(nFilasResultado, 6).Value = cant
            
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    
End Sub


Sub Listar_Pagos_Faltantes()
    Dim contenido As String
    Dim wbContenido As Workbook, _
        wsResultado As Excel.Worksheet, _
        wsContenido As Excel.Worksheet
    Dim nFilas As Double
    Dim nColumnas As Integer
    Dim rango As Range
    Dim nFilasCont As Double
    Dim rangoCont As Range
    Dim rangoError As Range
    Dim rangoTemp As String
    Dim i As Long
    Dim j As Long
    Dim m As Integer
    Dim valorJur As Integer
    Dim valorDoc As String
    Dim valorAct As String
    Dim resultado As Range
    Dim celdaDoc As String
    Dim temp As String
    Dim nFilasResultado As Integer
    Dim nFilasError As Integer
    Dim nColumnasError As Integer
    Dim bandera As Boolean
    Dim totalImporte As Double

    
    'Indicar el libro de excel CONTENIDO y control de errores
    contenido = InputBox("Ingrese el nombre del archivo:", "Abrir", "Archivo.xlsx")
    If contenido <> "" Then
        On Error GoTo ControlErrorOpen
        Set wbContenido = Workbooks.Open(ActiveWorkbook.Path & "\" & contenido)
    Else
        Exit Sub
    End If
    
    'Activar este libro
    ThisWorkbook.Activate
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Resultados"
    Application.DisplayAlerts = True
    
    Set wsResultado = Worksheets("Resultados")
    Set wsContenido = wbContenido.Worksheets("Detalle x Agente")
    
    'Regresa el control a la hoja de origen
    Sheets("Pagado").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    'Fila del encabezado Resultado
    nFilasResultado = 1
    wsResultado.Cells(nFilasResultado, 1).Value = "Jur"
    wsResultado.Cells(nFilasResultado, 2).Value = "Esc"
    wsResultado.Cells(nFilasResultado, 3).Value = "DNI"
    wsResultado.Cells(nFilasResultado, 4).Value = "Nombre y Apellido"
    wsResultado.Cells(nFilasResultado, 5).Value = "Importe Total"
    nFilasResultado = 2
    
    
    For i = 2 To nFilasCont
        valorDoc = wsContenido.Cells(i, 4).Value
        totalImporte = 0
        'Busca en el otro archivo
        rangoTemp = "B2:B" & nFilas
        Set resultado = ActiveSheet.Range(rangoTemp).Find(What:=valorDoc, _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False, _
                    SearchFormat:=False)
        'Si el resultado de la búsqueda no es vacío
        If Not resultado Is Nothing Then
            primerResultado = resultado.Address
            'Se obtiene el valor de j
            celdaDoc = resultado.Address
            tempDoc = ""
            For m = 1 To Len(celdaDoc)
                If IsNumeric(Mid(celdaDoc, m, 1)) Then
                    tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                End If
            Next m
            j = tempDoc
            
            cont = 0
            Do
                If wsContenido.Cells(i, 19).Value > 0 Then
                    totalImporte = totalImporte + wsContenido.Cells(i, 19).Value
                    cont = cont + 1
                End If
                i = i + 1
            Loop While valorDoc = wsContenido.Cells(i, 4).Value
            
            i = i - 1
            
            'cantidad de pagos < a
            'If Cells(j, 4).Value < 5 Then
            If Cells(j, 4).Value < cont And totalImporte > 0 Then
                wsResultado.Cells(nFilasResultado, 1).Value = wsContenido.Cells(i, 1).Value
                wsResultado.Cells(nFilasResultado, 2).Value = wsContenido.Cells(i, 7).Value
                wsResultado.Cells(nFilasResultado, 3).Value = wsContenido.Cells(i, 4).Value
                wsResultado.Cells(nFilasResultado, 4).Value = wsContenido.Cells(i, 6).Value
                wsResultado.Cells(nFilasResultado, 5).Value = totalImporte
                nFilasResultado = nFilasResultado + 1
            End If
        Else
            'No se encontró el documento
            Do
                totalImporte = totalImporte + wsContenido.Cells(i, 19).Value
                i = i + 1
            Loop While valorDoc = wsContenido.Cells(i, 4).Value
            
            i = i - 1
            If totalImporte > 0 Then
                wsResultado.Cells(nFilasResultado, 1).Value = wsContenido.Cells(i, 1).Value
                wsResultado.Cells(nFilasResultado, 2).Value = wsContenido.Cells(i, 7).Value
                wsResultado.Cells(nFilasResultado, 3).Value = wsContenido.Cells(i, 4).Value
                wsResultado.Cells(nFilasResultado, 4).Value = wsContenido.Cells(i, 6).Value
                wsResultado.Cells(nFilasResultado, 5).Value = totalImporte
                nFilasResultado = nFilasResultado + 1
            End If
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub



Sub Listar_Pagos()
    Dim contenido As String
    Dim wbContenido As Workbook, _
        wsResultado As Excel.Worksheet
    Dim nFilas As Double
    Dim nColumnas As Integer
    Dim rango As Range
    Dim i As Long
    Dim j As Long
    Dim valorDoc As String
    Dim resultado As Range
    Dim nFilasResultado As Integer
    Dim totalImporte As Double

    
    'Indicar el libro de excel CONTENIDO y control de errores
    contenido = InputBox("Ingrese el nombre del archivo:", "Abrir", "Archivo.xlsx")
    If contenido <> "" Then
        On Error GoTo ControlErrorOpen
        Set wbContenido = Workbooks.Open(ActiveWorkbook.Path & "\" & contenido)
    Else
        Exit Sub
    End If
    
    'Activar este libro
    ThisWorkbook.Activate
    
    Set wsResultado = wbContenido.Worksheets("Pagos 233")
    
    'Regresa el control a la hoja de origen
    Sheets("Detalle x Agente").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Fila del encabezado Resultado
    nFilasResultado = 1
    wsResultado.Cells(nFilasResultado, 1).Value = "Jur"
    wsResultado.Cells(nFilasResultado, 2).Value = "Esc"
    wsResultado.Cells(nFilasResultado, 3).Value = "DNI"
    wsResultado.Cells(nFilasResultado, 4).Value = "Nombre y Apellido"
    wsResultado.Cells(nFilasResultado, 5).Value = "Importe Total"
    wsResultado.Cells(nFilasResultado, 6).Value = "Cant Liq"
    wsResultado.Cells(nFilasResultado, 7).Value = "Total Liq"
    nFilasResultado = 2
    
    
    For i = 2 To nFilas
        valorDoc = Cells(i, 4).Value
        totalImporte = 0
        cont = 0
        cont2 = 0
        
        wsResultado.Cells(nFilasResultado, 1).Value = Cells(i, 1).Value
        wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 7).Value
        wsResultado.Cells(nFilasResultado, 3).Value = Cells(i, 4).Value
        wsResultado.Cells(nFilasResultado, 4).Value = Cells(i, 6).Value
                
        Do
            If Cells(i, 19).Value > 0 Then
                totalImporte = totalImporte + Cells(i, 19).Value
                cont = cont + 1
            End If
            cont2 = cont2 + 1
            i = i + 1
        Loop While valorDoc = Cells(i, 4).Value
        
        wsResultado.Cells(nFilasResultado, 5).Value = totalImporte
        wsResultado.Cells(nFilasResultado, 6).Value = cont
        wsResultado.Cells(nFilasResultado, 7).Value = cont2
        nFilasResultado = nFilasResultado + 1
        i = i - 1

    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub



Sub Listar_Comparacion()
    Dim contenido As String
    Dim wsPagado As Excel.Worksheet, _
        wsResultado As Excel.Worksheet
    Dim nFilas As Double
    Dim nColumnas As Integer
    Dim rango As Range
    Dim i As Long
    Dim j As Long
    Dim valorDoc As String
    Dim resultado As Range
    Dim nFilasResultado As Integer
    Dim totalImporte As Double

    Application.DisplayAlerts = False
    'Agrega las nuevas hojas
    Worksheets.Add
    ActiveSheet.Name = "Resultado"
    Application.DisplayAlerts = True
    
    Set wsResultado = Worksheets("Resultado")
    
    Set wsPagado = Worksheets("Pagados")
    
    'Regresa el control a la hoja de origen
    Sheets("Pagos 233").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsPagado.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    'Fila del encabezado Resultado
    nFilasResultado = 1
    wsResultado.Cells(nFilasResultado, 1).Value = "Jur"
    wsResultado.Cells(nFilasResultado, 2).Value = "Esc"
    wsResultado.Cells(nFilasResultado, 3).Value = "DNI"
    wsResultado.Cells(nFilasResultado, 4).Value = "Nombre y Apellido"
    wsResultado.Cells(nFilasResultado, 5).Value = "Importe Correspondiente"
    wsResultado.Cells(nFilasResultado, 6).Value = "Importe Liquidado"
    wsResultado.Cells(nFilasResultado, 7).Value = "Diferencia"
    nFilasResultado = 2
    
    
    For i = 2 To nFilas
        If Cells(i, 5).Value > 0 Then
            valorDoc = Cells(i, 3).Value
            'Busca en el otro archivo
            rangoTemp = "B2:B" & nFilasCont
            Set resultado = wsPagado.Range(rangoTemp).Find(What:=valorDoc, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False, _
                        SearchFormat:=False)
            'Si el resultado de la búsqueda no es vacío
            If Not resultado Is Nothing Then
                primerResultado = resultado.Address
                'Se obtiene el valor de j
                celdaDoc = resultado.Address
                tempDoc = ""
                For m = 1 To Len(celdaDoc)
                    If IsNumeric(Mid(celdaDoc, m, 1)) Then
                        tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                    End If
                Next m
                j = tempDoc
                
                If Abs(Cells(i, 6).Value - wsPagado.Cells(j, 4).Value) > 1 Then
                    wsResultado.Cells(nFilasResultado, 1).Value = Cells(i, 1).Value
                    wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 2).Value
                    wsResultado.Cells(nFilasResultado, 3).Value = Cells(i, 3).Value
                    wsResultado.Cells(nFilasResultado, 4).Value = Cells(i, 4).Value
                    wsResultado.Cells(nFilasResultado, 5).Value = Cells(i, 5).Value
                    
                    wsResultado.Cells(nFilasResultado, 6).Value = wsPagado.Cells(j, 5).Value
                    wsResultado.Cells(nFilasResultado, 7).Value = Cells(i, 5).Value - wsPagado.Cells(j, 5).Value
                    nFilasResultado = nFilasResultado + 1
                End If
            Else
                wsResultado.Cells(nFilasResultado, 1).Value = Cells(i, 1).Value
                wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 2).Value
                wsResultado.Cells(nFilasResultado, 3).Value = Cells(i, 3).Value
                wsResultado.Cells(nFilasResultado, 4).Value = Cells(i, 4).Value
                wsResultado.Cells(nFilasResultado, 5).Value = Cells(i, 5).Value
                wsResultado.Cells(nFilasResultado, 6).Value = 0
                wsResultado.Cells(nFilasResultado, 7).Value = Cells(i, 5).Value
                nFilasResultado = nFilasResultado + 1
            End If
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


Sub Listar_Liquidaciones()
    Dim contenido As String
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet
    Dim nFilas As Double
    Dim nColumnas As Integer
    Dim rango As Range
    Dim nFilasCont As Double
    Dim rangoCont As Range
    Dim rangoTemp As String
    Dim i As Long
    Dim j As Long
    Dim valorJur As Integer
    Dim valorDoc As String
    
    
    'Indicar el libro de excel CONTENIDO y control de errores
    contenido = InputBox("Ingrese el nombre del archivo:", "Abrir", "Archivo.xlsx")
    If contenido <> "" Then
        On Error GoTo ControlErrorOpen
        Set wbContenido = Workbooks.Open(ActiveWorkbook.Path & "\" & contenido)
    Else
        Exit Sub
    End If
    
    'Activar este libro
    ThisWorkbook.Activate
    
    Set wsContenido = wbContenido.Worksheets("Hoja1")
    
    'Regresa el control a la hoja de origen
    Sheets(1).Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    
    For i = 2 To nFilas
        valorDoc = Cells(i, 12).Value
        valorVto = Cells(i, 16).Value
        If Cells(i, 6).Value = 0 Then
            tempMes = Cells(i, 2).Value
            tempAnio = Cells(i, 1).Value
        Else
            For m = 1 To Len(valorVto)
                If Mid(valorVto, m, 1) = "/" Then
                    tempMes = Mid(valorVto, m + 1, 1)
                    m = m + 2
                    If Mid(valorVto, m, 1) <> "/" Then
                        tempMes = tempMes & Mid(valorVto, m, 1)
                        m = m + 1
                    End If
                    tempAnio = Mid(valorVto, m + 1, 4)
                    m = 10
                End If
            Next m
        End If
        
        'Busca en el otro archivo
        rangoTemp = "B2:B" & nFilasCont
        Set resultado = wsContenido.Range(rangoTemp).Find(What:=valorDoc, _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False, _
                    SearchFormat:=False)
        'Si el resultado de la búsqueda no es vacío
        If Not resultado Is Nothing Then
            primerResultado = resultado.Address
            'Se obtiene el valor de j
            celdaDoc = resultado.Address
            tempDoc = ""
            For m = 1 To Len(celdaDoc)
                If IsNumeric(Mid(celdaDoc, m, 1)) Then
                    tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                End If
            Next m
            j = tempDoc
            
            
            'retrocedo hasta encontrar el primero
            bandera = True
            Do While bandera
                If wsContenido.Cells(j - 1, 4).Value = valorDoc Then
                    j = j - 1
                Else
                    bandera = False
                End If
            Loop
            
            bandera = True
            Do
                If wsContenido.Cells(j, 5).Value = tempAnio And wsContenido.Cells(j, 6).Value = tempMes Then
                    bandera = False
                    'actualizar monto
                    If Cells(i, 6).Value = 2 Then
                        wsContenido.Cells(j, 9).Value = wsContenido.Cells(j, 9).Value - Cells(i, 7).Value
                    Else
                        wsContenido.Cells(j, 9).Value = wsContenido.Cells(j, 9).Value + Cells(i, 7).Value
                    End If
                    wsContenido.Cells(j, 10).Value = wsContenido.Cells(j, 10).Value + 1
                    wsContenido.Cells(j, 11).Value = wsContenido.Cells(j, 11).Value & " + " & Cells(i, 8).Value
                End If
                j = j + 1
            Loop While valorDoc = wsContenido.Cells(j, 2).Value And bandera
                
            If bandera Then
                'agregar
                wsContenido.Rows(j & ":" & j).Insert
                wsContenido.Cells(j, 1).Value = Cells(i, 8).Value
                wsContenido.Cells(j, 2).Value = valorDoc
                wsContenido.Cells(j, 3).Value = Cells(i, 14).Value
                wsContenido.Cells(j, 4).Value = Cells(i, 15).Value
                wsContenido.Cells(j, 5).Value = tempAnio
                wsContenido.Cells(j, 6).Value = tempMes
                wsContenido.Cells(j, 7).Value = Cells(i, 1).Value
                wsContenido.Cells(j, 8).Value = Cells(i, 2).Value
                If Cells(i, 6).Value = 2 Then
                    wsContenido.Cells(j, 9).Value = Cells(i, 7).Value * (-1)
                Else
                    wsContenido.Cells(j, 9).Value = Cells(i, 7).Value
                End If
                wsContenido.Cells(j, 10).Value = 1
                nFilasCont = nFilasCont + 1
            End If
            
        Else
            'No se encontró el documento
            nFilasCont = nFilasCont + 1
            wsContenido.Cells(nFilasCont, 1).Value = Cells(i, 8).Value
            wsContenido.Cells(nFilasCont, 2).Value = valorDoc
            wsContenido.Cells(nFilasCont, 3).Value = Cells(i, 14).Value
            wsContenido.Cells(nFilasCont, 4).Value = Cells(i, 15).Value
            wsContenido.Cells(nFilasCont, 5).Value = tempAnio
            wsContenido.Cells(nFilasCont, 6).Value = tempMes
            wsContenido.Cells(nFilasCont, 7).Value = Cells(i, 1).Value
            wsContenido.Cells(nFilasCont, 8).Value = Cells(i, 2).Value
            If Cells(i, 6).Value = 2 Then
                wsContenido.Cells(nFilasCont, 9).Value = Cells(i, 7).Value * (-1)
            Else
                wsContenido.Cells(nFilasCont, 9).Value = Cells(i, 7).Value
            End If
            wsContenido.Cells(nFilasCont, 10).Value = 1
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


Sub Listar_Liquidaciones_2()
    Dim contenido As String
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet
    Dim nFilas As Double
    Dim nColumnas As Integer
    Dim rango As Range
    Dim nFilasCont As Double
    Dim rangoCont As Range
    Dim rangoTemp As String
    Dim i As Long
    Dim j As Long
    Dim valorJur As Integer
    Dim valorDoc As String
    
    
    'Indicar el libro de excel CONTENIDO y control de errores
    contenido = InputBox("Ingrese el nombre del archivo:", "Abrir", "Archivo.xlsx")
    If contenido <> "" Then
        On Error GoTo ControlErrorOpen
        Set wbContenido = Workbooks.Open(ActiveWorkbook.Path & "\" & contenido)
    Else
        Exit Sub
    End If
    
    'Activar este libro
    ThisWorkbook.Activate
    
    Set wsContenido = wbContenido.Worksheets("Hoja1")
    
    'Regresa el control a la hoja de origen
    Sheets(1).Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    
    For i = 2 To nFilas
        valorDoc = Cells(i, 12).Value
        valorVto = Cells(i, 16).Value
        If Cells(i, 6).Value = 0 Then
            tempMes = Cells(i, 2).Value
            tempAnio = Cells(i, 1).Value
        Else
            For m = 1 To Len(valorVto)
                If Mid(valorVto, m, 1) = "/" Then
                    tempMes = Mid(valorVto, m + 1, 1)
                    m = m + 2
                    If Mid(valorVto, m, 1) <> "/" Then
                        tempMes = tempMes & Mid(valorVto, m, 1)
                        m = m + 1
                    End If
                    tempAnio = Mid(valorVto, m + 1, 4)
                    m = 10
                End If
            Next m
        End If
        
        'Busca en el otro archivo
        rangoTemp = "B2:B" & nFilasCont
        Set resultado = wsContenido.Range(rangoTemp).Find(What:=valorDoc, _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False, _
                    SearchFormat:=False)
        'Si el resultado de la búsqueda no es vacío
        If Not resultado Is Nothing Then
            primerResultado = resultado.Address
            'Se obtiene el valor de j
            celdaDoc = resultado.Address
            tempDoc = ""
            For m = 1 To Len(celdaDoc)
                If IsNumeric(Mid(celdaDoc, m, 1)) Then
                    tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                End If
            Next m
            j = tempDoc
            
            
            'retrocedo hasta encontrar el primero
            bandera = True
            Do While bandera
                If wsContenido.Cells(j - 1, 4).Value = valorDoc Then
                    j = j - 1
                Else
                    bandera = False
                End If
            Loop
            
            bandera = True
            Do
                If wsContenido.Cells(j, 5).Value = tempAnio And wsContenido.Cells(j, 6).Value = tempMes Then
                    bandera = False
                    'actualizar monto
                    If Cells(i, 6).Value = 2 Then
                        wsContenido.Cells(j, 9).Value = wsContenido.Cells(j, 9).Value - Cells(i, 7).Value
                    Else
                        wsContenido.Cells(j, 9).Value = wsContenido.Cells(j, 9).Value + Cells(i, 7).Value
                    End If
                    wsContenido.Cells(j, 10).Value = wsContenido.Cells(j, 10).Value + 1
                    wsContenido.Cells(j, 11).Value = wsContenido.Cells(j, 11).Value & " + " & Cells(i, 8).Value
                End If
                j = j + 1
            Loop While valorDoc = wsContenido.Cells(j, 2).Value And bandera
                
            If bandera Then
                'controlar al final
                j = nFilasCont
                bandera2 = True
                Do While valorDoc = wsContenido.Cells(j, 2).Value
                    If wsContenido.Cells(j, 5).Value = tempAnio And wsContenido.Cells(j, 6).Value = tempMes Then
                        bandera2 = False
                        If Cells(i, 6).Value = 2 Then
                            wsContenido.Cells(j, 9).Value = wsContenido.Cells(j, 9).Value - Cells(i, 7).Value
                        Else
                            wsContenido.Cells(j, 9).Value = wsContenido.Cells(j, 9).Value + Cells(i, 7).Value
                        End If
                        wsContenido.Cells(j, 10).Value = wsContenido.Cells(j, 10).Value + 1
                        wsContenido.Cells(j, 11).Value = wsContenido.Cells(j, 11).Value & " + " & Cells(i, 8).Value
                    End If
                    j = j - 1
                Loop
                If bandera2 Then
                    'agregar
                    nFilasCont = nFilasCont + 1
                    wsContenido.Cells(nFilasCont, 1).Value = Cells(i, 8).Value
                    wsContenido.Cells(nFilasCont, 2).Value = valorDoc
                    wsContenido.Cells(nFilasCont, 3).Value = Cells(i, 14).Value
                    wsContenido.Cells(nFilasCont, 4).Value = Cells(i, 15).Value
                    wsContenido.Cells(nFilasCont, 5).Value = tempAnio
                    wsContenido.Cells(nFilasCont, 6).Value = tempMes
                    wsContenido.Cells(nFilasCont, 7).Value = Cells(i, 1).Value
                    wsContenido.Cells(nFilasCont, 8).Value = Cells(i, 2).Value
                    If Cells(i, 6).Value = 2 Then
                        wsContenido.Cells(nFilasCont, 9).Value = Cells(i, 7).Value * (-1)
                    Else
                        wsContenido.Cells(nFilasCont, 9).Value = Cells(i, 7).Value
                    End If
                    wsContenido.Cells(nFilasCont, 10).Value = 1
                End If
            End If
            
        Else
            'No se encontró el documento
            nFilasCont = nFilasCont + 1
            wsContenido.Cells(nFilasCont, 1).Value = Cells(i, 8).Value
            wsContenido.Cells(nFilasCont, 2).Value = valorDoc
            wsContenido.Cells(nFilasCont, 3).Value = Cells(i, 14).Value
            wsContenido.Cells(nFilasCont, 4).Value = Cells(i, 15).Value
            wsContenido.Cells(nFilasCont, 5).Value = tempAnio
            wsContenido.Cells(nFilasCont, 6).Value = tempMes
            wsContenido.Cells(nFilasCont, 7).Value = Cells(i, 1).Value
            wsContenido.Cells(nFilasCont, 8).Value = Cells(i, 2).Value
            If Cells(i, 6).Value = 2 Then
                wsContenido.Cells(nFilasCont, 9).Value = Cells(i, 7).Value * (-1)
            Else
                wsContenido.Cells(nFilasCont, 9).Value = Cells(i, 7).Value
            End If
            wsContenido.Cells(nFilasCont, 10).Value = 1
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


Sub Controlar_Repetidos()
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim i As Long
    
    Sheets(1).Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    For i = 2 To nFilas
        If Cells(i, 2).Value = Cells(i + 1, 2).Value And Cells(i, 5).Value = Cells(i + 1, 5).Value And Cells(i, 6).Value = Cells(i + 1, 6).Value Then
            Cells(i, 9).Value = Cells(i, 9).Value + Cells(i + 1, 9).Value
            Cells(i, 10).Value = Cells(i, 10).Value + Cells(i + 1, 10).Value
            If Cells(i + 1, 11).Value <> "" Then
                Cells(i, 11).Value = Cells(i, 11).Value & " + " & Cells(i + 1, 11).Value
            End If
            nFilas = nFilas - 1
            Rows(i + 1).Delete
            i = i - 1
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Listar_Fechas()
    Dim contenido As String
    Dim wbContenido As Workbook, _
        wsResultado As Excel.Worksheet, _
        wsContenido As Excel.Worksheet
    Dim nFilas As Double
    Dim nColumnas As Integer
    Dim rango As Range
    Dim nFilasCont As Double
    Dim rangoCont As Range
    Dim i As Long
    Dim j As Long
    Dim m As Integer
    Dim valorDoc As String
    Dim resultado As Range
    Dim celdaDoc As String
    Dim temp As String
    Dim nFilasResultado As Integer
    Dim bandera As Boolean

    
    'Indicar el libro de excel CONTENIDO y control de errores
    contenido = InputBox("Ingrese el nombre del archivo:", "Abrir", "Archivo.xlsx")
    If contenido <> "" Then
        On Error GoTo ControlErrorOpen
        Set wbContenido = Workbooks.Open(ActiveWorkbook.Path & "\" & contenido)
    Else
        Exit Sub
    End If
    
    'Activar este libro
    ThisWorkbook.Activate
    
    'Borra las hojas destino si existen
    Application.DisplayAlerts = False
    
    'Agrega las nuevas hojas
    Worksheets.Add
    ActiveSheet.Name = "Informe"
    Application.DisplayAlerts = True
    
    Set wsResultado = Worksheets("Informe")
    Set wsContenido = wbContenido.Worksheets("Detalle x Agente")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    'Fila del encabezado Resultado
    nFilasResultado = 1
    wsResultado.Cells(nFilasResultado, 1).Value = "JUR"
    wsResultado.Cells(nFilasResultado, 2).Value = "DNI"
    wsResultado.Cells(nFilasResultado, 3).Value = "NOMBRE"
    wsResultado.Cells(nFilasResultado, 4).Value = "CEIC"
    wsResultado.Cells(nFilasResultado, 5).Value = "FEC_DESDE"
    wsResultado.Cells(nFilasResultado, 6).Value = "FEC_HASTA"
    wsResultado.Cells(nFilasResultado, 7).Value = "PAG_FEC_D"
    wsResultado.Cells(nFilasResultado, 8).Value = "PAG_FEC_H"
    wsResultado.Cells(nFilasResultado, 9).Value = "Observación"
    wsResultado.Cells(nFilasResultado, 10).Value = "Importe Agregar"
    
    valorDoc = "0"
    
    For i = 2 To nFilas
        If Cells(i, 9).Value > 0 Then
            If valorDoc <> Cells(i, 2).Value Then
                'If wsResultado.Cells(nFilasResultado, 8).Value <> "" And wsResultado.Cells(nFilasResultado, 10).Value <> "" Then
                If wsResultado.Cells(nFilasResultado, 10).Value = "" Then
                    wsResultado.Cells(nFilasResultado, 8).Value = Cells(i - 1, 6).Value & "-" & Cells(i - 1, 5).Value
                End If
                
                nFilasResultado = nFilasResultado + 1
                wsResultado.Cells(nFilasResultado, 1).Value = Cells(i, 1).Value
                wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 2).Value
                wsResultado.Cells(nFilasResultado, 3).Value = Cells(i, 3).Value
            
                valorDoc = Cells(i, 2).Value
                'Busca en el otro archivo
                rangoTemp = "D2:D" & nFilasCont
                Set resultado = wsContenido.Range(rangoTemp).Find(What:=valorDoc, _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False, _
                            SearchFormat:=False)
                'Si el resultado de la búsqueda no es vacío
                If Not resultado Is Nothing Then
                    primerResultado = resultado.Address
                    'Se obtiene el valor de j
                    celdaDoc = resultado.Address
                    tempDoc = ""
                    For m = 1 To Len(celdaDoc)
                        If IsNumeric(Mid(celdaDoc, m, 1)) Then
                            tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                        End If
                    Next m
                    j = tempDoc
                    
                    wsResultado.Cells(nFilasResultado, 4).Value = wsContenido.Cells(j, 8).Value
                    wsResultado.Cells(nFilasResultado, 5).Value = wsContenido.Cells(j, 10).Value
                    
                    Do
                        j = j + 1
                    Loop While valorDoc = wsContenido.Cells(j + 1, 4).Value
                    
                    wsResultado.Cells(nFilasResultado, 6).Value = wsContenido.Cells(j, 11).Value
                Else
                    'No se encontró el documento
                    wsResultado.Cells(nFilasResultado, 9).Value = "No se encontró DNI"
                    wsResultado.Cells(nFilasResultado, 4).Value = Cells(i, 4).Value
                End If
                'Controlar primer mes
                wsResultado.Cells(nFilasResultado, 7).Value = Cells(i, 6).Value & "-" & Cells(i, 5).Value
            Else
                'Controlar siguientes meses
                If Not (Cells(i, 6).Value - Cells(i - 1, 6).Value = 1 And Cells(i - 1, 5).Value = Cells(i, 5).Value And Cells(i, 9).Value > 0) Then
                    If Not (Cells(i - 1, 6).Value - Cells(i, 6).Value = 11 And Cells(i, 5).Value - Cells(i - 1, 5).Value = 1 And Cells(i, 9).Value > 0) Then
                        wsResultado.Cells(nFilasResultado, 8).Value = Cells(i - 1, 6).Value & "-" & Cells(i - 1, 5).Value
                        
                        nFilasResultado = nFilasResultado + 1
                        wsResultado.Cells(nFilasResultado, 1).Value = Cells(i, 1).Value
                        wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 2).Value
                        wsResultado.Cells(nFilasResultado, 3).Value = Cells(i, 3).Value
                        wsResultado.Cells(nFilasResultado, 4).Value = wsResultado.Cells(nFilasResultado - 1, 4).Value
                        wsResultado.Cells(nFilasResultado, 5).Value = wsResultado.Cells(nFilasResultado - 1, 5).Value
                        wsResultado.Cells(nFilasResultado, 6).Value = wsResultado.Cells(nFilasResultado - 1, 6).Value
                        wsResultado.Cells(nFilasResultado, 7).Value = Cells(i, 6).Value & "-" & Cells(i, 5).Value
                        wsResultado.Cells(nFilasResultado, 9).Value = wsResultado.Cells(nFilasResultado - 1, 9).Value
                    End If
                End If
            End If
        Else
            If Cells(i, 9).Value < 0 Then
                If valorDoc <> Cells(i, 2).Value Then
                    nFilasResultado = nFilasResultado + 1
                    wsResultado.Cells(nFilasResultado, 1).Value = Cells(i, 1).Value
                    wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 2).Value
                    wsResultado.Cells(nFilasResultado, 3).Value = Cells(i, 3).Value
                    wsResultado.Cells(nFilasResultado, 4).Value = Cells(i, 4).Value
                    valorDoc = Cells(i, 2).Value
                End If
                monto = -Cells(i, 9).Value
                'cont = 0
                j = i
                bandera = True
                Do While (monto - Cells(j - 1, 9).Value) > -1 And Cells(i, 2).Value = Cells(j - 1, 2).Value And bandera
                    'cont = cont + 1
                    j = j - 1
                    If Cells(j, 9).Value >= 0 Then
                        monto = monto - Cells(j, 9).Value
                    Else
                        bandera = False
                        j = j + 1
                    End If
                Loop
                'cont = cont + 1
                If Cells(j, 6).Value <> 1 Then
                    lk = (Cells(j, 6).Value)
                    efom = (Cells(j, 6).Value - 1)
                    wsResultado.Cells(nFilasResultado, 8).Value = (Cells(j, 6).Value - 1) & "-" & Cells(j, 5).Value
                Else
                    wsResultado.Cells(nFilasResultado, 8).Value = 12 & "-" & (Cells(j, 5).Value - 1)
                End If
                If monto > 1 Then
                    wsResultado.Cells(nFilasResultado, 10).Value = monto
                Else
                    wsResultado.Cells(nFilasResultado, 10).Value = 0
                End If
                valorDoc = "0"
                If wsResultado.Cells(nFilasResultado, 7).Value > wsResultado.Cells(nFilasResultado, 8).Value Then
                    wsResultado.Cells(nFilasResultado, 8).Value = ""
                    wsResultado.Cells(nFilasResultado, 7).Value = ""
                End If
            End If
        End If
    Next i
    
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub

Sub Calcular_Faltante()
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim i As Long
    Dim valorDoc As String
    Dim f_inicio As Date
    Dim f_desde As Date
    Dim f_pago As Date
    Dim f_ultima As Date
    
    Sheets("Informe").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    Cells(1, nColumnas + 1).Value = "Pagar F_Desde"
    Cells(1, nColumnas + 2).Value = "Pagar F_Hasta"
    Cells(1, nColumnas + 3).Value = "Desde - A la fecha"
    Cells(1, nColumnas + 4).Value = "Importe Agregar"
    
    valorDoc = "0"
    f_inicio = DateValue("Mar 1, 2015")
    f_ultima = DateValue("May 1, 2018")
    
    For i = 2 To nFilas
        If Cells(i, 5).Value <> "" Then
            If Cells(i, 7).Value <> "" Then
                If valorDoc <> Cells(i, 2).Value Then
                    valorDoc = Cells(i, 2).Value
                    f_desde = Cells(i, 5).Value
                    f_cobro = Cells(i, 7).Value
                    If f_cobro > f_desde And f_cobro > f_inicio Then
                        If f_desde > f_inicio Then
                            Cells(i, nColumnas + 1).Value = f_desde
                            Cells(i, nColumnas + 2).Value = DateAdd("m", -1, f_cobro)
                        Else
                            Cells(i, nColumnas + 1).Value = f_inicio
                            Cells(i, nColumnas + 2).Value = DateAdd("m", -1, f_cobro)
                        End If
                    End If
                    If Cells(i, 10).Value <> "" Then
                        Cells(i, nColumnas + 4).Value = Cells(i, 10).Value
                    End If
                Else
                    Cells(i, nColumnas + 1).Value = DateAdd("m", 1, Cells(i - 1, 8).Value)
                    Cells(i, nColumnas + 2).Value = DateAdd("m", -1, Cells(i, 7).Value)
                End If
                
                'controlar fecha de fin, en caso de ser un registro, o varios
                If Cells(i, 8).Value > f_ultima And Cells(i + 1, 2).Value <> valorDoc Then
                    Cells(i, nColumnas).Value = "CONTROLAR"
                End If
                If Cells(i, 8).Value < f_ultima And Cells(i + 1, 2).Value <> valorDoc Then
                    Cells(i, nColumnas + 3).Value = DateAdd("m", 1, Cells(i, 8).Value)
                End If
            Else
                If Cells(i, 8).Value = "" Then
                    Cells(i, nColumnas).Value = "CONTROLAR si corresponde"
                Else
                    If valorDoc <> Cells(i, 2).Value Then
                        'cobro antes
                        Cells(i, nColumnas + 1).Value = Cells(i, 8).Value
                        Cells(i, nColumnas + 4).Value = Cells(i, 10).Value
                    Else
                        'ver si cobra despues
                        bandera = False
                        m = i
                        Do While valorDoc = Cells(m + 1, 2).Value
                            m = m + 1
                            If Cells(m, 7).Value <> "" Then
                                bandera = True
                            End If
                        Loop
                        If bandera Then
                            Cells(i, nColumnas + 1).Value = Cells(i, 8).Value
                            Cells(i, nColumnas + 4).Value = Cells(i, 10).Value
                        End If
                    End If
                End If
            End If
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Calcular_Pagos()
    Dim contenido As String
    Dim wbContenido As Workbook, _
        wsResultado As Excel.Worksheet, _
        wsContenido As Excel.Worksheet
    Dim nFilas As Double
    Dim nColumnas As Integer
    Dim rango As Range
    Dim nFilasCont As Double
    Dim rangoCont As Range
    Dim i As Long
    Dim j As Long
    Dim nFilasResultado As Long
    Dim fecha As Date
    Dim ult_pago As Date
    Dim importe As Double

    
    'Indicar el libro de excel CONTENIDO y control de errores
    contenido = InputBox("Ingrese el nombre del archivo:", "Abrir", "Archivo.xlsx")
    If contenido <> "" Then
        On Error GoTo ControlErrorOpen
        Set wbContenido = Workbooks.Open(ActiveWorkbook.Path & "\" & contenido)
    Else
        Exit Sub
    End If
    
    'Activar este libro
    ThisWorkbook.Activate
    
    MsgBox "Debe estar ordenado por DNI y Fechas.", , "Atención"
    
    'Borra las hojas destino si existen
    Application.DisplayAlerts = False
    
    'Agrega las nuevas hojas
    Worksheets.Add
    ActiveSheet.Name = "Resultado"
    Application.DisplayAlerts = True
    
    Set wsResultado = Worksheets("Resultado")
    Set wsContenido = wbContenido.Worksheets("Totales")
    
    'Regresa el control a la hoja de origen
    Sheets("Informe").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    'Fila del encabezado Resultado
    nFilasResultado = 1
    wsResultado.Cells(nFilasResultado, 1).Value = "PtaId"
    wsResultado.Cells(nFilasResultado, 2).Value = "Jur"
    wsResultado.Cells(nFilasResultado, 3).Value = "EscId"
    wsResultado.Cells(nFilasResultado, 4).Value = "Pref"
    wsResultado.Cells(nFilasResultado, 5).Value = "Doc"
    wsResultado.Cells(nFilasResultado, 6).Value = "Digito"
    wsResultado.Cells(nFilasResultado, 7).Value = "Nombres"
    wsResultado.Cells(nFilasResultado, 8).Value = "Couc"
    wsResultado.Cells(nFilasResultado, 9).Value = "Reajuste"
    wsResultado.Cells(nFilasResultado, 10).Value = "Unidad"
    wsResultado.Cells(nFilasResultado, 11).Value = "Importe"
    wsResultado.Cells(nFilasResultado, 12).Value = "Vto"
    wsResultado.Cells(nFilasResultado, 13).Value = "Total"
    
    ult_pago = DateValue("Jun 1, 2018")
    For i = 2 To nFilas
        j = 2
        If Cells(i, 10).Value <> "" Then
            If Cells(i, 11).Value <> "" Then
                'Busca en el otro archivo
                Do While wsContenido.Cells(j, 1).Value <> Cells(i, 4).Value And j < nFilasCont
                    j = j + 1
                Loop
                
                If wsContenido.Cells(j, 1).Value = Cells(i, 4).Value Then
                    porc = 1
                    fecha = Cells(i, 10).Value
                    dia = Day(fecha)
                    If dia > 1 Then
                        'REGLA DE TRES SIMPLE CON RESPECTO AL DIA
                        dias = Day(DateAdd("m", 1, fecha - dia))
                        porc = dia / dias
                        fecha = fecha - dia + 1
                    End If
                    
                    importe = 0
                    Do
                        If fecha = wsContenido.Cells(j, 8).Value Then
                            nFilasResultado = nFilasResultado + 1
                            wsResultado.Cells(nFilasResultado, 1).Value = 0
                            wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 1).Value
                            wsResultado.Cells(nFilasResultado, 3).Value = 2
                            wsResultado.Cells(nFilasResultado, 4).Value = 0
                            wsResultado.Cells(nFilasResultado, 5).Value = Cells(i, 2).Value
                            wsResultado.Cells(nFilasResultado, 6).Value = 0
                            wsResultado.Cells(nFilasResultado, 7).Value = Cells(i, 3).Value
                            wsResultado.Cells(nFilasResultado, 8).Value = 233
                            wsResultado.Cells(nFilasResultado, 9).Value = 1
                            wsResultado.Cells(nFilasResultado, 10).Value = 0
                            wsResultado.Cells(nFilasResultado, 11).Value = (wsContenido.Cells(j, 4).Value) * porc
                            wsResultado.Cells(nFilasResultado, 12).Value = wsContenido.Cells(j, 5).Value
                            importe = importe + wsContenido.Cells(j, 4).Value
                            porc = 1
                            
                            'SAC
                            mes = wsContenido.Cells(j, 5).Value \ 10000
                            If mes = 6 Or mes = 12 Then
                                nFilasResultado = nFilasResultado + 1
                                wsResultado.Cells(nFilasResultado, 1).Value = 0
                                wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 1).Value
                                wsResultado.Cells(nFilasResultado, 3).Value = 2
                                wsResultado.Cells(nFilasResultado, 4).Value = 0
                                wsResultado.Cells(nFilasResultado, 5).Value = Cells(i, 2).Value
                                wsResultado.Cells(nFilasResultado, 6).Value = 0
                                wsResultado.Cells(nFilasResultado, 7).Value = Cells(i, 3).Value
                                wsResultado.Cells(nFilasResultado, 8).Value = 316
                                wsResultado.Cells(nFilasResultado, 9).Value = 1
                                wsResultado.Cells(nFilasResultado, 10).Value = 0
                                wsResultado.Cells(nFilasResultado, 11).Value = wsResultado.Cells(nFilasResultado - 1, 11).Value / 2
                                wsResultado.Cells(nFilasResultado, 12).Value = wsContenido.Cells(j, 5).Value
                                importe = importe + wsContenido.Cells(j, 4).Value / 2
                            End If
                            
                            fecha = DateAdd("m", 1, fecha)
                        End If
                        j = j + 1
                    Loop While Cells(i, 11).Value >= fecha
                    wsResultado.Cells(nFilasResultado, 13).Value = importe
                    
                    If Cells(i, 10).Value = "" And Cells(i - 1, 2).Value = Cells(i, 2).Value And Cells(i - 1, 13).Value > 0 Then
                        'importe agregado
                        nFilasResultado = nFilasResultado + 1
                        wsResultado.Cells(nFilasResultado, 1).Value = 0
                        wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 1).Value
                        wsResultado.Cells(nFilasResultado, 3).Value = 2
                        wsResultado.Cells(nFilasResultado, 4).Value = 0
                        wsResultado.Cells(nFilasResultado, 5).Value = Cells(i, 2).Value
                        wsResultado.Cells(nFilasResultado, 6).Value = 0
                        wsResultado.Cells(nFilasResultado, 7).Value = Cells(i, 3).Value
                        wsResultado.Cells(nFilasResultado, 8).Value = 233
                        wsResultado.Cells(nFilasResultado, 9).Value = 1
                        wsResultado.Cells(nFilasResultado, 10).Value = 0
                        wsResultado.Cells(nFilasResultado, 11).Value = Cells(i - 1, 13).Value
                        tempFecha = Cells(i, 10).Value
                        wsResultado.Cells(nFilasResultado, 12).Value = Month(tempFecha) & Year(tempFecha)
                    End If

                    If Cells(i, 13).Value > 0 And Cells(i, 12).Value = "" Then
                        'importe agregado
                        nFilasResultado = nFilasResultado + 1
                        wsResultado.Cells(nFilasResultado, 1).Value = 0
                        wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 1).Value
                        wsResultado.Cells(nFilasResultado, 3).Value = 2
                        wsResultado.Cells(nFilasResultado, 4).Value = 0
                        wsResultado.Cells(nFilasResultado, 5).Value = Cells(i, 2).Value
                        wsResultado.Cells(nFilasResultado, 6).Value = 0
                        wsResultado.Cells(nFilasResultado, 7).Value = Cells(i, 3).Value
                        wsResultado.Cells(nFilasResultado, 8).Value = 233
                        wsResultado.Cells(nFilasResultado, 9).Value = 1
                        wsResultado.Cells(nFilasResultado, 10).Value = 0
                        wsResultado.Cells(nFilasResultado, 11).Value = Cells(i, 13).Value
                        tempFecha = Cells(i, 11).Value
                        wsResultado.Cells(nFilasResultado, 12).Value = Month(tempFecha) & Year(tempFecha)
                    End If
                Else
                    Cells(i, 9).Value = "No se encontró CEIC"
                End If
            Else
                If Cells(i, 12).Value = "" Then
                    If Cells(i, 13).Value > 0 Then
                        'importe agregado
                        nFilasResultado = nFilasResultado + 1
                        wsResultado.Cells(nFilasResultado, 1).Value = 0
                        wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 1).Value
                        wsResultado.Cells(nFilasResultado, 3).Value = 2
                        wsResultado.Cells(nFilasResultado, 4).Value = 0
                        wsResultado.Cells(nFilasResultado, 5).Value = Cells(i, 2).Value
                        wsResultado.Cells(nFilasResultado, 6).Value = 0
                        wsResultado.Cells(nFilasResultado, 7).Value = Cells(i, 3).Value
                        wsResultado.Cells(nFilasResultado, 8).Value = 233
                        wsResultado.Cells(nFilasResultado, 9).Value = 1
                        wsResultado.Cells(nFilasResultado, 10).Value = 0
                        wsResultado.Cells(nFilasResultado, 11).Value = Cells(i, 13).Value
                        tempFecha = Cells(i, 10).Value
                        wsResultado.Cells(nFilasResultado, 12).Value = Month(tempFecha) & Year(tempFecha)
                    End If
                End If
            End If
        End If
        
        If Cells(i, 12).Value <> "" Then
            'repetir lo anterior desde la fecha Cel-12 hasta ult_pago
            'Busca en el otro archivo
            Do While wsContenido.Cells(j, 1).Value <> Cells(i, 4).Value And j < nFilasCont
                j = j + 1
            Loop
            
            If wsContenido.Cells(j, 1).Value = Cells(i, 4).Value Then
                fecha = Cells(i, 12).Value
                importe = 0
                Do
                    If fecha = wsContenido.Cells(j, 8).Value Then
                        nFilasResultado = nFilasResultado + 1
                        wsResultado.Cells(nFilasResultado, 1).Value = 0
                        wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 1).Value
                        wsResultado.Cells(nFilasResultado, 3).Value = 2
                        wsResultado.Cells(nFilasResultado, 4).Value = 0
                        wsResultado.Cells(nFilasResultado, 5).Value = Cells(i, 2).Value
                        wsResultado.Cells(nFilasResultado, 6).Value = 0
                        wsResultado.Cells(nFilasResultado, 7).Value = Cells(i, 3).Value
                        wsResultado.Cells(nFilasResultado, 8).Value = 233
                        wsResultado.Cells(nFilasResultado, 9).Value = 1
                        wsResultado.Cells(nFilasResultado, 10).Value = 0
                        wsResultado.Cells(nFilasResultado, 11).Value = wsContenido.Cells(j, 4).Value
                        wsResultado.Cells(nFilasResultado, 12).Value = wsContenido.Cells(j, 5).Value
                        importe = importe + wsContenido.Cells(j, 4).Value
                        
                        'SAC
                        mes = wsContenido.Cells(j, 5).Value \ 10000
                        If mes = 6 Or mes = 12 Then
                            nFilasResultado = nFilasResultado + 1
                            wsResultado.Cells(nFilasResultado, 1).Value = 0
                            wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 1).Value
                            wsResultado.Cells(nFilasResultado, 3).Value = 2
                            wsResultado.Cells(nFilasResultado, 4).Value = 0
                            wsResultado.Cells(nFilasResultado, 5).Value = Cells(i, 2).Value
                            wsResultado.Cells(nFilasResultado, 6).Value = 0
                            wsResultado.Cells(nFilasResultado, 7).Value = Cells(i, 3).Value
                            wsResultado.Cells(nFilasResultado, 8).Value = 316
                            wsResultado.Cells(nFilasResultado, 9).Value = 1
                            wsResultado.Cells(nFilasResultado, 10).Value = 0
                            wsResultado.Cells(nFilasResultado, 11).Value = wsContenido.Cells(j, 4).Value / 2
                            wsResultado.Cells(nFilasResultado, 12).Value = wsContenido.Cells(j, 5).Value
                            importe = importe + wsContenido.Cells(j, 4).Value / 2
                        End If
                        
                        fecha = DateAdd("m", 1, fecha)
                    End If
                    j = j + 1
                Loop While ult_pago >= fecha
                wsResultado.Cells(nFilasResultado, 13).Value = importe
            Else
                Cells(i, 9).Value = "No se encontró CEIC"
            End If
            
            If Cells(i, 13).Value > 0 Then
                'importe agregado
                nFilasResultado = nFilasResultado + 1
                wsResultado.Cells(nFilasResultado, 1).Value = 0
                wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 1).Value
                wsResultado.Cells(nFilasResultado, 3).Value = 2
                wsResultado.Cells(nFilasResultado, 4).Value = 0
                wsResultado.Cells(nFilasResultado, 5).Value = Cells(i, 2).Value
                wsResultado.Cells(nFilasResultado, 6).Value = 0
                wsResultado.Cells(nFilasResultado, 7).Value = Cells(i, 3).Value
                wsResultado.Cells(nFilasResultado, 8).Value = 233
                wsResultado.Cells(nFilasResultado, 9).Value = 1
                wsResultado.Cells(nFilasResultado, 10).Value = 0
                wsResultado.Cells(nFilasResultado, 11).Value = Cells(i, 13).Value
                tempFecha = Cells(i, 12).Value
                wsResultado.Cells(nFilasResultado, 12).Value = Month(tempFecha) & Year(tempFecha)
            End If
        End If
    Next i
    
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub

Sub Calcular_Pagos_Nuevo()
    Dim contenido As String
    Dim wbContenido As Workbook, _
        wsResultado As Excel.Worksheet, _
        wsContenido As Excel.Worksheet
    Dim nFilas As Double
    Dim nColumnas As Integer
    Dim rango As Range
    Dim nFilasCont As Double
    Dim rangoCont As Range
    Dim i As Long
    Dim j As Long
    Dim nFilasResultado As Long
    Dim fecha As Date
    Dim ult_pago As Date
    Dim importe As Double

    
    'Indicar el libro de excel CONTENIDO y control de errores
    contenido = InputBox("Ingrese el nombre del archivo:", "Abrir", "Archivo.xlsx")
    If contenido <> "" Then
        On Error GoTo ControlErrorOpen
        Set wbContenido = Workbooks.Open(ActiveWorkbook.Path & "\" & contenido)
    Else
        Exit Sub
    End If
    
    'Activar este libro
    ThisWorkbook.Activate
    
    MsgBox "Debe estar ordenado por DNI y Fechas.", , "Atención"
    
    'Borra las hojas destino si existen
    Application.DisplayAlerts = False
    
    'Agrega las nuevas hojas
    Worksheets.Add
    ActiveSheet.Name = "Resultado"
    Application.DisplayAlerts = True
    
    Set wsResultado = Worksheets("Resultado")
    Set wsContenido = wbContenido.Worksheets("Totales")
    
    'Regresa el control a la hoja de origen
    Sheets("Informe").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    'Fila del encabezado Resultado
    nFilasResultado = 1
    wsResultado.Cells(nFilasResultado, 1).Value = "PtaId"
    wsResultado.Cells(nFilasResultado, 2).Value = "Jur"
    wsResultado.Cells(nFilasResultado, 3).Value = "EscId"
    wsResultado.Cells(nFilasResultado, 4).Value = "Pref"
    wsResultado.Cells(nFilasResultado, 5).Value = "Doc"
    wsResultado.Cells(nFilasResultado, 6).Value = "Digito"
    wsResultado.Cells(nFilasResultado, 7).Value = "Nombres"
    wsResultado.Cells(nFilasResultado, 8).Value = "Couc"
    wsResultado.Cells(nFilasResultado, 9).Value = "Reajuste"
    wsResultado.Cells(nFilasResultado, 10).Value = "Unidad"
    wsResultado.Cells(nFilasResultado, 11).Value = "Importe"
    wsResultado.Cells(nFilasResultado, 12).Value = "Vto"
    wsResultado.Cells(nFilasResultado, 13).Value = "Total"
    
    ult_pago = DateValue("Jun 1, 2018")
    
    For i = 2 To nFilas
        If Month(Cells(i, 8).Value) = Month(ult_pago) And Year(Cells(i, 8).Value) = Year(ult_pago) Then
            j = i
            Do
                Cells(j, nColumnas + 1).Value = "Pagado"
                j = j - 1
            Loop While Cells(i, 2).Value = Cells(j, 2).Value
        End If
    Next i
    
    For i = 2 To nFilas
        j = 2
        If Cells(i, nColumnas + 1).Value = "Pagado" Then
            If Cells(i, 10).Value <> "" And Cells(i, 11).Value <> "" Then
                'Busca en el otro archivo
                Do While wsContenido.Cells(j, 1).Value <> Cells(i, 4).Value And j < nFilasCont
                    j = j + 1
                Loop
                
                If wsContenido.Cells(j, 1).Value = Cells(i, 4).Value Then
                    porc = 1
                    fecha = Cells(i, 10).Value
                    dia = Day(fecha)
                    If dia > 1 Then
                        'REGLA DE TRES SIMPLE CON RESPECTO AL DIA
                        dias = Day(DateAdd("m", 1, fecha - dia))
                        porc = dia / dias
                        fecha = fecha - dia + 1
                    End If
                    
                    importe = 0
                    Do
                        If Month(fecha) = wsContenido.Cells(j, 6).Value And Year(fecha) = wsContenido.Cells(j, 7).Value Then
                            nFilasResultado = nFilasResultado + 1
                            wsResultado.Cells(nFilasResultado, 1).Value = 0
                            wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 1).Value
                            wsResultado.Cells(nFilasResultado, 3).Value = 2
                            wsResultado.Cells(nFilasResultado, 4).Value = 0
                            wsResultado.Cells(nFilasResultado, 5).Value = Cells(i, 2).Value
                            wsResultado.Cells(nFilasResultado, 6).Value = 0
                            wsResultado.Cells(nFilasResultado, 7).Value = Cells(i, 3).Value
                            wsResultado.Cells(nFilasResultado, 8).Value = 233
                            wsResultado.Cells(nFilasResultado, 9).Value = 1
                            wsResultado.Cells(nFilasResultado, 10).Value = 0
                            wsResultado.Cells(nFilasResultado, 11).Value = (wsContenido.Cells(j, 4).Value) * porc
                            wsResultado.Cells(nFilasResultado, 12).Value = wsContenido.Cells(j, 5).Value
                            importe = importe + wsContenido.Cells(j, 4).Value
                            porc = 1
                            
                            'SAC
                            mes = wsContenido.Cells(j, 5).Value \ 10000
                            If mes = 6 Or mes = 12 Then
                                nFilasResultado = nFilasResultado + 1
                                wsResultado.Cells(nFilasResultado, 1).Value = 0
                                wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 1).Value
                                wsResultado.Cells(nFilasResultado, 3).Value = 2
                                wsResultado.Cells(nFilasResultado, 4).Value = 0
                                wsResultado.Cells(nFilasResultado, 5).Value = Cells(i, 2).Value
                                wsResultado.Cells(nFilasResultado, 6).Value = 0
                                wsResultado.Cells(nFilasResultado, 7).Value = Cells(i, 3).Value
                                wsResultado.Cells(nFilasResultado, 8).Value = 316
                                wsResultado.Cells(nFilasResultado, 9).Value = 1
                                wsResultado.Cells(nFilasResultado, 10).Value = 0
                                wsResultado.Cells(nFilasResultado, 11).Value = wsResultado.Cells(nFilasResultado - 1, 11).Value / 2
                                wsResultado.Cells(nFilasResultado, 12).Value = wsContenido.Cells(j, 5).Value
                                importe = importe + wsContenido.Cells(j, 4).Value / 2
                            End If
                            
                            fecha = DateAdd("m", 1, fecha)
                        End If
                        j = j + 1
                    Loop While Cells(i, 11).Value >= fecha
                    
                    If Cells(i, 13).Value > 0 Then
                        'importe agregado
                        nFilasResultado = nFilasResultado + 1
                        wsResultado.Cells(nFilasResultado, 1).Value = 0
                        wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 1).Value
                        wsResultado.Cells(nFilasResultado, 3).Value = 2
                        wsResultado.Cells(nFilasResultado, 4).Value = 0
                        wsResultado.Cells(nFilasResultado, 5).Value = Cells(i, 2).Value
                        wsResultado.Cells(nFilasResultado, 6).Value = 0
                        wsResultado.Cells(nFilasResultado, 7).Value = Cells(i, 3).Value
                        wsResultado.Cells(nFilasResultado, 8).Value = 233
                        wsResultado.Cells(nFilasResultado, 9).Value = 1
                        wsResultado.Cells(nFilasResultado, 10).Value = 0
                        wsResultado.Cells(nFilasResultado, 11).Value = Cells(i, 13).Value
                        tempFecha = Cells(i, 11).Value
                        wsResultado.Cells(nFilasResultado, 12).Value = Month(tempFecha) & Year(tempFecha)
                        importe = importe + Cells(i, 13).Value
                    End If
                    
                    wsResultado.Cells(nFilasResultado, 13).Value = importe
                Else
                    Cells(i, 9).Value = "No se encontró CEIC"
                End If
            Else
                If Cells(i, 13).Value > 0 Then
                    If Cells(i, 10).Value <> "" Then
                        nFilasResultado = nFilasResultado + 1
                        wsResultado.Cells(nFilasResultado, 1).Value = 0
                        wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 1).Value
                        wsResultado.Cells(nFilasResultado, 3).Value = 2
                        wsResultado.Cells(nFilasResultado, 4).Value = 0
                        wsResultado.Cells(nFilasResultado, 5).Value = Cells(i, 2).Value
                        wsResultado.Cells(nFilasResultado, 6).Value = 0
                        wsResultado.Cells(nFilasResultado, 7).Value = Cells(i, 3).Value
                        wsResultado.Cells(nFilasResultado, 8).Value = 233
                        wsResultado.Cells(nFilasResultado, 9).Value = 1
                        wsResultado.Cells(nFilasResultado, 10).Value = 0
                        tempFecha = Cells(i, 10).Value
                        wsResultado.Cells(nFilasResultado, 11).Value = Cells(i, 13).Value
                        wsResultado.Cells(nFilasResultado, 12).Value = Month(tempFecha) & Year(tempFecha)
                    Else
                        If Cells(i, 10).Value = "" And Cells(i + 1, 2).Value = Cells(i, 2).Value Then
                            nFilasResultado = nFilasResultado + 1
                            wsResultado.Cells(nFilasResultado, 1).Value = 0
                            wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 1).Value
                            wsResultado.Cells(nFilasResultado, 3).Value = 2
                            wsResultado.Cells(nFilasResultado, 4).Value = 0
                            wsResultado.Cells(nFilasResultado, 5).Value = Cells(i, 2).Value
                            wsResultado.Cells(nFilasResultado, 6).Value = 0
                            wsResultado.Cells(nFilasResultado, 7).Value = Cells(i, 3).Value
                            wsResultado.Cells(nFilasResultado, 8).Value = 233
                            wsResultado.Cells(nFilasResultado, 9).Value = 1
                            wsResultado.Cells(nFilasResultado, 10).Value = 0
                            wsResultado.Cells(nFilasResultado, 11).Value = Cells(i, 13).Value
                            tempFecha = Cells(i, 10).Value
                            wsResultado.Cells(nFilasResultado, 12).Value = Month(fecha) & Year(fecha)
                        End If
                    End If
                End If
            End If
        End If
    Next i
    
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


Sub Buscar_Fechas()
    Dim contenido As String
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet
    Dim nFilas As Double
    Dim nColumnas As Integer
    Dim rango As Range
    Dim nFilasCont As Double
    Dim rangoCont As Range
    Dim i As Long
    Dim valorDoc As String
    Dim resultado As Range
    Dim celdaDoc As String
    Dim temp As String
    Dim fecha1 As Date
    Dim fecha2 As Date
    
    'Indicar el libro de excel CONTENIDO y control de errores
    contenido = InputBox("Ingrese el nombre del archivo:", "Abrir", "Archivo.xlsx")
    If contenido <> "" Then
        On Error GoTo ControlErrorOpen
        Set wbContenido = Workbooks.Open(ActiveWorkbook.Path & "\" & contenido)
    Else
        Exit Sub
    End If
    
    'Activar este libro
    ThisWorkbook.Activate
    
    'Borra las hojas destino si existen
    Application.DisplayAlerts = False
    
    Set wsContenido = wbContenido.Worksheets(1)
    
    'Regresa el control a la hoja de origen
    Sheets("Informe").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    
    For i = 2 To nFilas
        If Cells(i, 5).Value = "" Then
            valorDoc = Cells(i, 2).Value
            
            'Busca en el otro archivo
            rangoTemp = "C2:C" & nFilasCont
            Set resultado = wsContenido.Range(rangoTemp).Find(What:=valorDoc, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False, _
                        SearchFormat:=False)
            'Si el resultado de la búsqueda no es vacío
            If Not resultado Is Nothing Then
                primerResultado = resultado.Address
                'Se obtiene el valor de j
                celdaDoc = resultado.Address
                tempDoc = ""
                For m = 1 To Len(celdaDoc)
                    If IsNumeric(Mid(celdaDoc, m, 1)) Then
                        tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                    End If
                Next m
                j = tempDoc
                
                Cells(i, 5).Value = wsContenido.Cells(j, 4).Value
                Cells(i, 6).Value = wsContenido.Cells(j, 23).Value
                Cells(i, 9).Value = ""
                
                Do While valorDoc = wsContenido.Cells(j + 1, 3).Value
                    fecha1 = wsContenido.Cells(j + 1, 4).Value
                    fecha2 = wsContenido.Cells(j + 1, 23).Value
                    If Cells(i, 5).Value > fecha1 Then
                        Cells(i, 5).Value = fecha1
                    End If
                    If Cells(i, 6).Value < fecha2 Then
                        Cells(i, 6).Value = fecha2
                    End If
                    j = j + 1
                Loop
            End If
        End If
    Next i
    
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub

Sub Actualizar_Pagos()
    Dim contenido As String
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet
    Dim wsInforme As Excel.Worksheet
    Dim nFilas As Double
    Dim nColumnas As Integer
    Dim rango As Range
    Dim nFilasCont As Double
    Dim nFilasInf As Double
    Dim rangoCont As Range
    Dim rangoTemp As String
    Dim i As Long
    Dim j As Long
    Dim valorVto As Date
    Dim valorDoc As String
    
    MsgBox "Debe estar ordenado por DNI y Vto.", , "Atención!!"
    
    'Indicar el libro de excel CONTENIDO y control de errores
    contenido = InputBox("Ingrese el nombre del archivo:", "Abrir", "Archivo.xlsx")
    If contenido <> "" Then
        On Error GoTo ControlErrorOpen
        Set wbContenido = Workbooks.Open(ActiveWorkbook.Path & "\" & contenido)
    Else
        Exit Sub
    End If
    
    'Activar este libro
    ThisWorkbook.Activate
    
    Set wsContenido = wbContenido.Worksheets("Hoja1")
    Set wsInforme = wbContenido.Worksheets("Informe")
    
    'Regresa el control a la hoja de origen
    Sheets(1).Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    Set rangoCont2 = wsInforme.UsedRange
    nFilasInf = rangoCont2.Rows.Count
    
    primerFila = nFilasCont + 1
    
    For i = 2 To nFilas
        valorVto = Cells(i, 16).Value
        tempMes = Month(valorVto)
        tempAnio = Year(valorVto)
                
        'Agregar en Hoja1
        nFilasCont = nFilasCont + 1
        wsContenido.Cells(nFilasCont, 1).Value = Cells(i, 8).Value
        wsContenido.Cells(nFilasCont, 2).Value = Cells(i, 12).Value
        wsContenido.Cells(nFilasCont, 3).Value = Cells(i, 14).Value
        wsContenido.Cells(nFilasCont, 4).Value = Cells(i, 15).Value
        wsContenido.Cells(nFilasCont, 5).Value = tempAnio
        wsContenido.Cells(nFilasCont, 6).Value = tempMes
        wsContenido.Cells(nFilasCont, 7).Value = Cells(i, 1).Value
        wsContenido.Cells(nFilasCont, 8).Value = Cells(i, 2).Value
        wsContenido.Cells(nFilasCont, 9).Value = Cells(i, 7).Value
        wsContenido.Cells(nFilasCont, 10).Value = 1
    Next i
    


    For i = primerFila To nFilasCont
        valorDoc = wsContenido.Cells(i, 2).Value
        
        'Busca en el otro archivo
        rangoTemp = "B2:B" & nFilasInf
        Set resultado = wsInforme.Range(rangoTemp).Find(What:=valorDoc, _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False, _
                    SearchFormat:=False)
        'Si el resultado de la búsqueda no es vacío
        If Not resultado Is Nothing Then
            primerResultado = resultado.Address
            'Se obtiene el valor de j
            celdaDoc = resultado.Address
            tempDoc = ""
            For m = 1 To Len(celdaDoc)
                If IsNumeric(Mid(celdaDoc, m, 1)) Then
                    tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                End If
            Next m
            j = tempDoc
            
            tempMes = wsContenido.Cells(i, 6).Value
            tempAnio = wsContenido.Cells(i, 5).Value
            copCol = 15
            wsInforme.Cells(j, 14).Value = tempMes & "-" & tempAnio
            
            If valorDoc = wsContenido.Cells(i + 1, 2).Value Then
                i = i + 1
                Do
                    'hacer algo para que en las columnas 14-15 poner lo que se pagó
                    'para luego poder hacer una diferencia de fechas y ver si falta pagar algo
                    '14 fecha desde. 15 fecha hasta. 16 fecha desde y de ultima observar si hay algo raro
                    mes = tempMes
                    anio = tempAnio
                    tempMes = wsContenido.Cells(i, 6).Value
                    tempAnio = wsContenido.Cells(i, 5).Value
                    If (mes = tempMes - 1 And anio = tempAnio) Or (mes = 12 And tempMes = 1 And anio = tempAnio - 1) Then
                        i = i + 1
                    Else
                        wsInforme.Cells(j, copCol).Value = mes & "-" & anio
                        wsInforme.Cells(j, copCol + 1).Value = tempMes & "-" & tempAnio
                        copCol = copCol + 2
                        i = i + 1
                    End If
                Loop While valorDoc = wsContenido.Cells(i, 2).Value
                i = i - 1
            Else
                wsInforme.Cells(j, 15).Value = tempMes & "-" & tempAnio
            End If
            
            
        Else
            'No se encontró el documento
            nFilasInf = nFilasInf + 1
            wsInforme.Cells(nFilasInf, 1).Value = wsContenido.Cells(i, 1).Value
            wsInforme.Cells(nFilasInf, 2).Value = wsContenido.Cells(i, 2).Value
            wsInforme.Cells(nFilasInf, 3).Value = wsContenido.Cells(i, 3).Value
            wsInforme.Cells(nFilasInf, 4).Value = wsContenido.Cells(i, 4).Value
            wsInforme.Cells(nFilasInf, 7).Value = wsContenido.Cells(i, 6).Value & "-" & wsContenido.Cells(i, 5).Value
            
            Do While valorDoc = wsContenido.Cells(i + 1, 2).Value
                If wsContenido.Cells(i + 1, 6).Value - 1 = wsContenido.Cells(i, 6).Value And wsContenido.Cells(i + 1, 5).Value = wsContenido.Cells(i, 5).Value Then
                    i = i + 1
                Else
                    If wsContenido.Cells(i, 6).Value = 12 And wsContenido.Cells(i + 1, 5).Value = wsContenido.Cells(i, 5).Value - 1 And wsContenido.Cells(i + 1, 6).Value = 1 Then
                        i = i + 1
                    Else
                        valorDoc = ""
                    End If
                End If
            Loop
            wsInforme.Cells(nFilasInf, 8).Value = wsContenido.Cells(i, 6).Value & "-" & wsContenido.Cells(i, 5).Value
            wsInforme.Cells(nFilasInf, 9).Value = "Controlar"
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


Sub Actualizar_Pagos_Parte2()
    Dim contenido As String
    Dim nFilas As Double
    Dim nColumnas As Integer
    Dim rango As Range
    Dim rangoTemp As String
    Dim i As Long
    Dim j As Long
    Dim valorVto As Date
    Dim valorDoc As String
    
    Sheets("Informe").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    For i = 2 To nFilas
        If Cells(i, 14).Value <> "" Then
            If Cells(i, 14).Value = "1/6/2018" Then
                If Cells(i + 1, 2).Value = Cells(i, 2).Value Then
                    If Cells(i + 2, 2).Value = Cells(i, 2).Value Then
                        Cells(i + 2, 8).Value = "6-18"
                    Else
                        Cells(i + 1, 8).Value = "6-18"
                    End If
                Else
                    Cells(i, 8).Value = "6-18"
                End If
            Else
                If Cells(i, 14).Value = Cells(i, 10).Value And Cells(i, 15).Value = Cells(i, 11).Value Then
                    Cells(i, 7).Value = Month(Cells(i, 14).Value) & "-" & Year(Cells(i, 14).Value)
                    Cells(i, 10).Value = ""
                    Cells(i, 11).Value = ""
                    If Cells(i, 16).Value = "1/6/2018" Then
                        If Cells(i + 1, 2).Value = Cells(i, 2).Value Then
                            Cells(i + 1, 8).Value = "6-18"
                        Else
                            Cells(i, 8).Value = "6-18"
                        End If
                    Else
                        Cells(i, nColumnas + 1).Value = "CONTROLAR"
                    End If
                Else
                    If Cells(i, 14).Value = Cells(i + 1, 10).Value And Cells(i, 15).Value = Cells(i + 1, 11).Value And Cells(i + 1, 2).Value = Cells(i, 2).Value Then
                        If Cells(i, 16).Value = "1/6/2018" Then
                            Cells(i, 8).Value = "6-18"
                            'eliminar fila i+1
                        Else
                            Cells(i, 8).Value = Cells(i + 1, 8).Value
                            'eliminar fila i+1
                        End If
                        If Cells(i + 1, 14).Value <> "" Then
                            Cells(i, 14).Value = Cells(i + 1, 14).Value
                        End If
                        Rows(i + 1).Delete
                        nFilas = nFilas - 1
                    Else
                        Cells(i, nColumnas + 1).Value = "CONTROLAR"
                    End If
                End If
            End If
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    
End Sub


