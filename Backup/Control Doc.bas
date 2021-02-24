Attribute VB_Name = "Módulo2"
Sub Controlar_Documentos()
    Dim j As Long
    Dim rango As Range
    Dim rangoError As Range
    Dim resultado As Range
    Dim resultadoError As Range
    Dim nFilas As Long
    Dim nFilasError As Integer
    Dim nColumnasError As Integer
    Dim wsError As Excel.Worksheet
    Dim temp As String
    Dim rangoDocError As String
    Dim rangoDoc As String
    
    'Verificar si existen documentos repetidos en el archivo
    
    Application.DisplayAlerts = False
    'Sheets("Errores").Delete
    'Agrega la nueva hoja
    Worksheets.Add
    ActiveSheet.Name = "Errores"
    Application.DisplayAlerts = True

    Set wsError = Worksheets("Errores")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
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
    
    For j = 2 To nFilas
        'Primero revisa si el documento ya está en Errores (de esta forma no aparecerá varias veces el mismo documento en Errores)
        valorDoc = Cells(j, 5).Value
        rangoDocError = "E1:E" & nFilasError
        Set resultadoError = wsError.Range(rangoDocError).Find(What:=valorDoc, _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False, _
                    SearchFormat:=False)
        If resultadoError Is Nothing Then
            'Controla cuantas veces se repite el documento
            rangoDoc = "E" & j & ":E" & nFilas
            Set resultado = Range(rangoDoc).Find(What:=valorDoc, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False, _
                        SearchFormat:=False)
            cont = 0
            primerResultado = resultado.Address
            Do
                cont = cont + 1
                Set resultado = Range(rangoDoc).FindNext(resultado)
            Loop While Not resultado Is Nothing And primerResultado <> resultado.Address
            If cont > 1 Then
                temp = j & ":" & j
                Range(temp).Copy
                temp = nFilasError & ":" & nFilasError
                wsError.Range(temp).PasteSpecial xlPasteAll
                Application.CutCopyMode = False
                wsError.Cells(nFilasError, nColumnasError).Value = "Aparece " & cont & " veces en el archivo."
                nFilasError = nFilasError + 1
            End If
        End If
    Next j
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Eliminar_Documentos()
    Dim i As Long
    Dim rango As Range
    Dim rangoError As Range
    Dim rangoCont As Range
    Dim resultado As Range
    Dim resultadoError As Range
    Dim nFilas As Double
    Dim nFilasCont As Double
    Dim nFilasError As Integer
    Dim nColumnasError As Integer
    Dim wbContenido As Workbook, _
        wsError As Excel.Worksheet, _
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
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Errores"
    Application.DisplayAlerts = True

    Set wsError = Worksheets("Errores")
    Set wsContenido = wbContenido.Worksheets("A___HRG___Selec_vs_cptos_x_Juri")
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select

    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    'Fila del encabezado Error
    wsError.Cells(1, 1).Value = "JurId"
    wsError.Cells(1, 2).Value = "Esc"
    wsError.Cells(1, 3).Value = "Doc"
    wsError.Cells(1, 4).Value = "Nombre"
    wsError.Cells(1, 5).Value = "Mensaje"
    wsError.Columns(5).ColumnWidth = 52
    nFilasError = 2
    
    For i = 2 To nFilas
        valorDoc = Cells(i, 5).Value
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
            
            'Eliminar la fila y actualizar número
            Rows(j).Delete
            nFilasCont = nFilasCont - 1

        Else
            'No se encontró el documento
            wsError.Cells(nFilasError, 1).Value = Cells(i, 2).Value
            wsError.Cells(nFilasError, 2).Value = Cells(i, 3).Value
            wsError.Cells(nFilasError, 3).Value = Cells(i, 5).Value
            wsError.Cells(nFilasError, 4).Value = Cells(i, 7).Value
            wsError.Cells(nFilasError, 5).Value = "No se encontró el documento."
            nFilasError = nFilasError + 1
            
            Cells(i, nColumnas + 1).Value = "Ver en Errores"
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


Sub Copiar_Monto_Doc()
    Dim i As Long
    Dim rango As Range
    Dim rangoError As Range
    Dim rangoCont As Range
    Dim resultado As Range
    Dim resultadoError As Range
    Dim nFilas As Double
    Dim nFilasCont As Double
    Dim nFilasError As Integer
    Dim nColumnasError As Integer
    Dim wbContenido As Workbook, _
        wsError As Excel.Worksheet, _
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
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Errores"
    Application.DisplayAlerts = True

    Set wsError = Worksheets("Errores")
    Set wsContenido = wbContenido.Worksheets("A___HRG___Selec_vs_cptos_x_Juri")
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select

    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    'Fila del encabezado Error
    wsError.Cells(1, 1).Value = "JurId"
    wsError.Cells(1, 2).Value = "Esc"
    wsError.Cells(1, 3).Value = "Doc"
    wsError.Cells(1, 4).Value = "Nombre"
    wsError.Cells(1, 5).Value = "Mensaje"
    wsError.Columns(5).ColumnWidth = 52
    nFilasError = 2
    
    For i = 2 To nFilas
        valorDoc = Cells(i, 5).Value
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
            
            'Copiar el monto en el otro documento
            wsContenido.Cells(j, 20).Value = Cells(i, 12).Value / 2

        Else
            'No se encontró el documento
            wsError.Cells(nFilasError, 1).Value = Cells(i, 2).Value
            wsError.Cells(nFilasError, 2).Value = Cells(i, 3).Value
            wsError.Cells(nFilasError, 3).Value = Cells(i, 5).Value
            wsError.Cells(nFilasError, 4).Value = Cells(i, 7).Value
            wsError.Cells(nFilasError, 5).Value = "No se encontró el documento."
            nFilasError = nFilasError + 1
            
            Cells(i, nColumnas + 1).Value = "Ver en Errores"
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


Sub Control_Iguales()
    Dim i As Long
    Dim rango As Range
    Dim rangoError As Range
    Dim rangoCont As Range
    Dim resultado As Range
    Dim resultadoError As Range
    Dim nFilas As Double
    Dim nFilasCont As Double
    Dim nFilasError As Integer
    Dim nColumnasError As Integer
    Dim wbContenido As Workbook, _
        wsError As Excel.Worksheet, _
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
    
    Set wsContenido = wbContenido.Worksheets("Resultados")
    'Regresa el control a la hoja de origen
    Sheets("Resultados").Select

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
            
            Do While wsContenido.Cells(j - 1, 5).Value = valorDoc
                j = j - 1
            Loop
            
            Do While (wsContenido.Cells(j, 5).Value = valorDoc And wsContenido.Cells(j, 12).Value <> Cells(i, 12).Value)
                j = j + 1
            Loop
            
            If (wsContenido.Cells(j, 5).Value = valorDoc And wsContenido.Cells(j, 12).Value = Cells(i, 12).Value) Then
                If wsContenido.Cells(j, 8).Value <> Cells(i, 8).Value Then
                    j = j + 1
                End If
                For m = 1 To 12
                    bandera = False
                    If Cells(i, m).Value <> wsContenido.Cells(j, m).Value Then
                        bandera = True
                    End If
                    If bandera Then
                        Cells(i, nColumnas + 1).Value = "Modificado"
                    End If
                Next m
            Else
                Cells(i, nColumnas + 1).Value = "No encontrado"
            End If

        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub

