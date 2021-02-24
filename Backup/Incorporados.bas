Attribute VB_Name = "Módulo1"
Sub Acumular_Archivo()
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet
    Dim nFilas As Double
    Dim nColumnas As Integer
    Dim rango As Range
    Dim nFilasCont As Double
    Dim rangoCont As Range
    
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
    
    Sheets(1).Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    For i = 2 To nFilasCont
        temp = i & ":" & i
        wsContenido.Range(temp).Copy
        nFilas = nFilas + 1
        temp = nFilas & ":" & nFilas
        Range(temp).PasteSpecial xlPasteAll
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


Sub Calcular_JurMov()
    Dim nFilas As Long
    Dim nFilasCont As Long
    Dim nColumnas As Integer
    Dim wsResultado As Excel.Worksheet
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Resultados"
    Application.DisplayAlerts = True
    
    Set wsResultado = Worksheets("Resultados")
    
    'Regresa el control a la hoja de origen
    Sheets("Archivo Incorporados").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    wsResultado.Cells(1, 1).Value = "Total de Registros:"
    wsResultado.Cells(1, 2).Value = 0
    wsResultado.Cells(2, 1).Value = "JUR"
    wsResultado.Cells(2, 2).Value = "Cant de Registros"
    nFilasCont = 2
    
    For i = 2 To nFilas
        wsResultado.Cells(1, 2).Value = wsResultado.Cells(1, 2).Value + 1
        
        jur = Cells(i, 2).Value
        'Busca en el otro archivo
        rangoTemp = "A3:A" & nFilasCont
        Set resultado = wsResultado.Range(rangoTemp).Find(What:=jur, _
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
            
            wsResultado.Cells(j, 2).Value = wsResultado.Cells(j, 2).Value + 1
        Else
            nFilasCont = nFilasCont + 1
            wsResultado.Cells(nFilasCont, 1).Value = jur
            wsResultado.Cells(nFilasCont, 2).Value = 1
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Informe_Movimientos()
    Dim nFilas As Long
    Dim nFilasCont As Long
    Dim nColumnas As Integer
    Dim wsResultado As Excel.Worksheet
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Resultados"
    Application.DisplayAlerts = True
    
    Set wsResultado = Worksheets("Resultados")
    
    'Regresa el control a la hoja de origen
    Sheets("Movimientos Ajustes Autorizados").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    wsResultado.Cells(1, 1).Value = "Jur"
    wsResultado.Cells(1, 2).Value = "Actuación"
    wsResultado.Cells(1, 3).Value = "Fch.Autorizado"
    wsResultado.Cells(1, 4).Value = "Estado"
    wsResultado.Cells(1, 5).Value = "Complemento"
    wsResultado.Cells(1, 6).Value = "Operador"
    wsResultado.Cells(1, 7).Value = "Cantidad"
    wsResultado.Cells(1, 8).Value = "Importe Autor."
    wsResultado.Cells(1, 9).Value = "Importe Regis."
    wsResultado.Cells(1, 10).Value = "Importe Total"
    nFilasCont = 1
    
    For i = 2 To nFilas
        jur = Cells(i, 17).Value
        'Busca en el otro archivo
        rangoTemp = "B3:B" & nFilasCont
        Set resultado = wsResultado.Range(rangoTemp).Find(What:=jur, _
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
            
            If wsResultado.Cells(j, 4).Value <> Cells(i, 21).Value Then
                wsResultado.Cells(j, 4).Value = "Controlar"
            End If
            If Cells(i, 21).Value = "Registrado" Then
                wsResultado.Cells(j, 9).Value = Cells(i, 24).Value + wsResultado.Cells(j, 9).Value
            Else
                wsResultado.Cells(j, 8).Value = Cells(i, 24).Value + wsResultado.Cells(j, 8).Value
            End If
            wsResultado.Cells(j, 7).Value = wsResultado.Cells(j, 7).Value + 1
            wsResultado.Cells(j, 10).Value = Cells(i, 24).Value + wsResultado.Cells(j, 10).Value
        Else
            nFilasCont = nFilasCont + 1
            wsResultado.Cells(nFilasCont, 1).Value = Cells(i, 1).Value
            wsResultado.Cells(nFilasCont, 2).Value = Cells(i, 17).Value
            wsResultado.Cells(nFilasCont, 3).Value = Cells(i, 19).Value
            wsResultado.Cells(nFilasCont, 4).Value = Cells(i, 21).Value
            wsResultado.Cells(nFilasCont, 5).Value = Cells(i, 9).Value
            wsResultado.Cells(nFilasCont, 6).Value = Cells(i, 20).Value
            wsResultado.Cells(nFilasCont, 7).Value = 1
            wsResultado.Cells(nFilasCont, 10).Value = Cells(i, 24).Value
            If Cells(i, 21).Value = "Registrado" Then
                wsResultado.Cells(nFilasCont, 9).Value = Cells(i, 24).Value
            Else
                wsResultado.Cells(nFilasCont, 8).Value = Cells(i, 24).Value
            End If
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub


