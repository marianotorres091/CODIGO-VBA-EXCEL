Attribute VB_Name = "Módulo11"
Sub Controlar_Documentos()
    Dim i As Long
    Dim rango As Range
    Dim rangoCont As Range
    Dim resultado As Range
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim nFilasCont As Double
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet
    Dim temp As String
    
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
    
    
    For i = 2 To nFilas
        valorDoc = Cells(i, 5).Value
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
            
            Cells(i, nColumnas + 1).Value = wsContenido.Cells(j, 18).Value
            Cells(i, nColumnas + 2).Value = wsContenido.Cells(j, 14).Value
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub

Sub Filtrar_Con_Extras()
    Dim i As Long
    Dim nFilas As Double
    Dim nFilasResultado As Double
    Dim nColumnas As Double
    Dim rango As Range
    Dim wsResultado As Excel.Worksheet
    
    
    'Agrego la nueva hoja
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Ajuste 120"
    Application.DisplayAlerts = True
    Set wsResultado = Worksheets("Ajuste 120")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    Range("1:1").Copy
    wsResultado.Range("1:1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    nFilasResultado = 2
    
    For i = 2 To nFilas
        If Cells(i, nColumnas).Value <> "" Or Cells(i, nColumnas - 1).Value <> "" Then
            temp = i & ":" & i
            Range(temp).Copy
            temp = nFilasResultado & ":" & nFilasResultado
            wsResultado.Range(temp).PasteSpecial xlPasteAll
            Application.CutCopyMode = False
            nFilasResultado = nFilasResultado + 1
        End If
    Next i
End Sub


Sub Calcular_Totales()
    Dim i As Long
    Dim valorDoc As String
    Dim rango As Range
    Dim rangoCont As Range
    Dim resultado As Range
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim nFilasCont As Double
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet
    Dim temp As String
    Dim total As Double
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
    
    Cells(1, 17).Value = "Horas Extras"
    valorDoc = Cells(2, 2).Value
    total = 0
    bandera = False
    For i = 2 To nFilas
        If Cells(i, 2).Value = valorDoc Then
            If Cells(i, 10).Value = 2 Then
                total = total - Cells(i, 12).Value
            Else
                total = total + Cells(i, 12).Value
            End If
        Else
            Rows(i).Insert
            nFilas = nFilas + 1
            
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
                
                Cells(i, 1).Value = wsContenido.Cells(j, 2).Value
                Cells(i, 2).Value = wsContenido.Cells(j, 5).Value
                Cells(i, 3).Value = wsContenido.Cells(j, 7).Value
                Cells(i, 8).Value = wsContenido.Cells(j, 8).Value
                Cells(i, 10).Value = wsContenido.Cells(j, 9).Value
                Cells(i, 11).Value = wsContenido.Cells(j, 10).Value
                Cells(i, 12).Value = wsContenido.Cells(j, 11).Value
                Cells(i, 17).Value = wsContenido.Cells(j, 13).Value
                
                bandera = True
                total = total - Cells(i, 12).Value
                i = i + 1
                Rows(i).Insert
                nFilas = nFilas + 1
            End If
            Cells(i, 12).Value = total
            valorDoc = Cells(i + 1, 2).Value
            If bandera Then
                Cells(i, 17).Value = (total / 130) * Cells(i - 1, 17).Value
            End If
            total = 0
            bandera = False
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub

Sub Generar_Diferenia()
    Dim i As Long
    Dim valorDoc As String
    Dim rango As Range
    Dim rangoCont As Range
    Dim resultado As Range
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim nFilasCont As Double
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet, _
        wsResultado As Excel.Worksheet
    Dim filaResultado As Integer
    
    
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
    
    'Agrego la nueva hoja
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Resultado"
    Application.DisplayAlerts = True
    Set wsResultado = Worksheets("Resultado")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select

    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    wsResultado.Cells(1, 1).Value = "JurId"
    wsResultado.Cells(1, 2).Value = "Doc"
    wsResultado.Cells(1, 3).Value = "Nombre"
    wsResultado.Cells(1, 4).Value = "Horas Extras"
    wsResultado.Cells(1, 5).Value = "Importe Calculado"
    wsResultado.Cells(1, 6).Value = "Importe Recibido"
    wsResultado.Cells(1, 7).Value = "Diferencia"
    filaResultado = 2
    
    For i = 2 To nFilas
        If Cells(i, 17).Value <> "" Then
            valorDoc = Cells(i, 2).Value
            'Busca en el otro archivo
            rangoTemp = "L2:L" & nFilasCont
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
                
                importeRecibido = wsContenido.Cells(j, 7).Value
                
                wsResultado.Cells(filaResultado, 1).Value = Cells(i, 1).Value
                wsResultado.Cells(filaResultado, 2).Value = Cells(i, 2).Value
                wsResultado.Cells(filaResultado, 3).Value = Cells(i, 3).Value
                wsResultado.Cells(filaResultado, 4).Value = Cells(i, 17).Value
                wsResultado.Cells(filaResultado, 5).Value = Cells(i + 1, 17).Value
                wsResultado.Cells(filaResultado, 6).Value = importeRecibido
                wsResultado.Cells(filaResultado, 7).Value = Cells(i + 1, 17).Value - importeRecibido
                filaResultado = filaResultado + 1
                i = i + 1
            Else
                wsResultado.Cells(filaResultado, 1).Value = Cells(i, 1).Value
                wsResultado.Cells(filaResultado, 2).Value = Cells(i, 2).Value
                wsResultado.Cells(filaResultado, 3).Value = Cells(i, 3).Value
                wsResultado.Cells(filaResultado, 4).Value = Cells(i, 17).Value
                wsResultado.Cells(filaResultado, 5).Value = Cells(i + 1, 17).Value
                wsResultado.Cells(filaResultado, 6).Value = "No se encontró el documento"
                filaResultado = filaResultado + 1
                i = i + 1
            End If
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub
