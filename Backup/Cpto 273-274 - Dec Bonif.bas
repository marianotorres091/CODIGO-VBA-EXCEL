Attribute VB_Name = "Module1"
Sub Agregar_Importe()
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
    Dim valorDoc As String
    Dim resultado As Range
    Dim celdaDoc As String
    Dim temp As String
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
    
    Set wsContenido = wbContenido.Worksheets("A___HRG___Seleccion_de_Concepto")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    Cells(1, nColumnas + 1).Value = "Importe"
    Cells(1, nColumnas + 2).Value = "Importe Básico"
    Cells(1, nColumnas + 3).Value = "Observación"
    Cells(1, nColumnas + 4).Value = "CEIC"
    
    For i = 2 To nFilas
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
            
            Do While valorDoc = wsContenido.Cells(j - 1, 12).Value
                j = j - 1
            Loop
            
            bandera = True
            
            'controlar nombre
            'If Mid(wsContenido.Cells(j, 6).Value, 1, 3) <> Mid(Cells(i, 4).Value, 1, 3) Then
            '    Cells(i, nColumnas + 2).Value = "Controlar DNI+Nombre"
            'End If
            Cells(i, nColumnas + 4).Value = wsContenido.Cells(j, 15).Value
            
            If wsContenido.Cells(j, 4).Value = 1 Then
                Cells(i, nColumnas + 2).Value = wsContenido.Cells(j, 7).Value
                j = j + 1
            End If
            If wsContenido.Cells(j, 4).Value = 120 And Cells(i, 2).Value = wsContenido.Cells(j, 12).Value Then
                bandera = False
                Cells(i, nColumnas + 3).Value = "Cobra Cpto 120"
                j = j + 1
            End If
            If wsContenido.Cells(j, 4).Value = 126 And Cells(i, 2).Value = wsContenido.Cells(j, 12).Value Then
                If bandera Then
                    Cells(i, nColumnas + 3).Value = "Cobra Cpto 126"
                    bandera = False
                Else
                    Cells(i, nColumnas + 3).Value = Cells(i, nColumnas + 3).Value & " - 126"
                End If
                j = j + 1
            End If
            If wsContenido.Cells(j, 4).Value = 273 And Cells(i, 2).Value = wsContenido.Cells(j, 12).Value Then
                If Cells(i, 7).Value = wsContenido.Cells(j, 4).Value Then
                    If Cells(i, nColumnas + 3).Value = "" Then
                        Cells(i, nColumnas + 1).Value = wsContenido.Cells(j, 7).Value
                    End If
                Else
                    If bandera Then
                        Cells(i, nColumnas + 3).Value = "Cobra Cpto 273"
                        bandera = False
                    Else
                        Cells(i, nColumnas + 3).Value = Cells(i, nColumnas + 3).Value & " - 273"
                    End If
                    j = j + 1
                End If
            End If
            If wsContenido.Cells(j, 4).Value = 274 And Cells(i, 2).Value = wsContenido.Cells(j, 12).Value Then
                If Cells(i, 7).Value = wsContenido.Cells(j, 4).Value Then
                    If Cells(i, nColumnas + 3).Value = "" Then
                        Cells(i, nColumnas + 1).Value = wsContenido.Cells(j, 7).Value
                    End If
                Else
                    If bandera Then
                        Cells(i, nColumnas + 3).Value = "Cobra Cpto 274"
                    Else
                        Cells(i, nColumnas + 3).Value = Cells(i, nColumnas + 3).Value & " - 274"
                    End If
                    Cells(i, nColumnas + 1).Value = ""
                    j = j + 1
                End If
            End If
        Else
            'No se encontró el documento
            Cells(i, nColumnas + 3).Value = "No se encontró el DNI"
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub

Sub Calcular_Importe()
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
    Dim valorDoc As String
    Dim resultado As Range
    Dim celdaDoc As String
    Dim temp As String
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
    
    Set wsContenido = wbContenido.Worksheets(1)
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    Cells(1, nColumnas + 1).Value = "Importe Calculado"
    Cells(1, nColumnas + 2).Value = "Observación"
    
    For i = 2 To nFilas
        valorDoc = Cells(i, 13).Value
        'Busca en el otro archivo
        rangoTemp = "L3:L" & nFilasCont
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
            
            If Cells(i, 12).Value = "" Then
                If Cells(i, 7).Value = 273 Then
                    Cells(i, nColumnas + 1).Value = wsContenido.Cells(j, 10).Value * Cells(i, 3).Value
                Else
                    If Cells(i, 7).Value = 274 Then
                        If wsContenido.Cells(j, 9).Value > 0 Then
                            Cells(i, nColumnas + 1).Value = wsContenido.Cells(j, 9).Value * Cells(i, 3).Value
                        Else
                            col = 27 - wsContenido.Cells(j, 3).Value
                            Cells(i, nColumnas + 1).Value = wsContenido.Cells(col, 9).Value * Cells(i, 3).Value
                        End If
                    Else
                        Cells(i, nColumnas + 2).Value = "Error"
                    End If
                End If
            Else
                Cells(i, nColumnas + 2).Value = Cells(i, 12).Value
            End If
        Else
            'No se encontró el documento
            Cells(i, nColumnas + 2).Value = "No se encontró CEIC"
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub
