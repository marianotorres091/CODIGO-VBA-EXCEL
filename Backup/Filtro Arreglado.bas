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
    'Sheets("Resultados").Delete
    'Sheets("Errores").Delete
    
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
    wsResultado.Cells(nFilasResultado, 14).Value = "Actuación"
    nFilasResultado = 2
    
    Cells(1, nColumnas + 1).Value = "Observación"
    
    For i = 2 To nFilas
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
                        wsResultado.Cells(nFilasResultado, 14).Value = valorAct
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
                    
                    Cells(i, nColumnas + 1).Value = "Ver en Errores"
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
                
                Cells(i, nColumnas + 1).Value = "Ver en Errores"
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
            
            Cells(i, nColumnas + 1).Value = "Ver en Errores"
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
                    Cells(i, nColumnas + 1).Value = wsContenido.Cells(j, 3).Value + Cells(i, nColumnas + 1).Value
                    If wsContenido.Cells(j, 7).Value = 2 Then
                        Cells(i, nColumnas + 2).Value = Cells(i, nColumnas + 2).Value - wsContenido.Cells(j, 6).Value
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

