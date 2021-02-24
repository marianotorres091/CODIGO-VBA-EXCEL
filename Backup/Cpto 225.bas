Attribute VB_Name = "Módulo1"
Sub Comparar_Cargos()
    Dim i As Long
    Dim rango As Range
    Dim rangoCont As Range
    Dim resultado As Range
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim nColumnasCont As Double
    Dim nFilasCont As Double
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet
    Dim temp As String
    Dim cImporte As Integer
    Dim cPorcentaje As Integer
    Dim cCeic As Integer
    Dim bonificacion As Double
    Dim tempImporte As Double
    
    
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

    Set wsContenido = wbContenido.Worksheets("EscalaActual")
    'Regresa el control a la hoja de origen
    Sheets(1).Select

    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    nColumnasCont = rangoCont.Columns.Count
    
    cImporte = 7
    cPorcentaje = 18
    cCeic = 15
    
    For i = 2 To nFilas
        If Cells(i, cPorcentaje).Value > 0 Then
            bonificacion = Cells(i, cImporte).Value * 100 / Cells(i, cPorcentaje).Value
        
            'Busca en el otro archivo
            rangoTemp = "D2:D" & nFilasCont
            Set resultado = wsContenido.Range(rangoTemp).Find(What:=Cells(i, cCeic).Value, _
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
                
                If bonificacion > (wsContenido.Cells(j, nColumnasCont).Value + 10) Then
                    'Voy para arriba
                    For j = j + 1 To 3 Step -1
                        may = wsContenido.Cells(j, nColumnasCont).Value + 10
                        men = wsContenido.Cells(j, nColumnasCont).Value - 10
                        If bonificacion < may And bonificacion > men Then
                            Cells(i, nColumnas + 1).Value = wsContenido.Cells(j, 5).Value
                            Cells(i, nColumnas + 2).Value = wsContenido.Cells(j, 6).Value
                            'Cells(i, nColumnas + 3).Value = "Controlar. Grupo Superior"
                            j = 2
                        End If
                    Next
                    If Cells(i, nColumnas + 2).Value = "" Then
                        Cells(i, nColumnas + 3).Value = "Controlar. No se encontró grupo"
                    End If
                Else
                    If bonificacion < (wsContenido.Cells(j, nColumnasCont).Value - 10) Then
                        'Voy para abajo
                        For m = j - 1 To j - 3
                            may = wsContenido.Cells(m, nColumnasCont).Value + 10
                            men = wsContenido.Cells(m, nColumnasCont).Value - 10
                            If bonificacion < may And bonificacion > men Then
                                Cells(i, nColumnas + 1).Value = wsContenido.Cells(m, 5).Value
                                Cells(i, nColumnas + 2).Value = wsContenido.Cells(m, 6).Value
                                Cells(i, nColumnas + 3).Value = "Controlar. Grupo Inferior"
                                j = 2
                            End If
                        Next
                        If Cells(i, nColumnas + 3).Value = "" Then
                            Cells(i, nColumnas + 3).Value = "Controlar. No se encontró grupo"
                        End If
                    End If
                End If
            End If
        Else
            Cells(i, nColumnas + 3).Value = "Error Unidad"
        End If
    Next
    
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


Sub Comparar()
    Dim i As Long
    Dim rango As Range
    Dim rangoCont As Range
    Dim resultado As Range
    Dim nFilas As Double
    Dim nColumnas As Double
    
    
    'Activar este libro
    ThisWorkbook.Activate
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    For i = 2 To nFilas
        If (Asc(Cells(i, nColumnas - 6).Value) < Asc(Cells(i, nColumnas - 3).Value)) Or (Cells(i, nColumnas - 4).Value > Cells(i, nColumnas - 2).Value) Then
            Cells(i, nColumnas + 1).Value = "Controlar"
        End If
    Next
    
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
End Sub


Sub Filtrar()
    Dim i As Long
    Dim nFilasCont As Long
    Dim rango As Range
    Dim rangoCont As Range
    Dim resultado As Range
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim wsContenido As Excel.Worksheet
    
    'Activar este libro
    ThisWorkbook.Activate
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Resultado"
    Application.DisplayAlerts = True

    Set wsContenido = Worksheets("Resultado")
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    nFilasCont = 1
    wsContenido.Cells(nFilasCont, 1).Value = "DNI"
    wsContenido.Cells(nFilasCont, 2).Value = "NOMBRE"
    wsContenido.Cells(nFilasCont, 3).Value = "JUR"
    wsContenido.Cells(nFilasCont, 4).Value = "CEIC"
    wsContenido.Cells(nFilasCont, 5).Value = "IMPORTE"
    wsContenido.Cells(nFilasCont, 6).Value = "PORCENTAJE"
    wsContenido.Cells(nFilasCont, 7).Value = "APARTADO"
    wsContenido.Cells(nFilasCont, 8).Value = "GRUPO"
    wsContenido.Cells(nFilasCont, 9).Value = "AP_Calculado"
    wsContenido.Cells(nFilasCont, 10).Value = "GR_Calculado"
    wsContenido.Cells(nFilasCont, 11).Value = "Observación"
    nFilasCont = 2
    
    For i = 2 To nFilas
        If Cells(i, nColumnas).Value <> "" Or Cells(i, nColumnas - 1).Value <> "" Then
            wsContenido.Cells(nFilasCont, 1).Value = Cells(i, 12).Value
            wsContenido.Cells(nFilasCont, 2).Value = Cells(i, 14).Value
            wsContenido.Cells(nFilasCont, 3).Value = Cells(i, 8).Value
            wsContenido.Cells(nFilasCont, 4).Value = Cells(i, 15).Value
            wsContenido.Cells(nFilasCont, 5).Value = Cells(i, 7).Value
            wsContenido.Cells(nFilasCont, 6).Value = Cells(i, 18).Value
            wsContenido.Cells(nFilasCont, 7).Value = Cells(i, 24).Value
            wsContenido.Cells(nFilasCont, 8).Value = Cells(i, 26).Value
            wsContenido.Cells(nFilasCont, 9).Value = Cells(i, 27).Value
            wsContenido.Cells(nFilasCont, 10).Value = Cells(i, 28).Value
            wsContenido.Cells(nFilasCont, 11).Value = Cells(i, 29).Value
            nFilasCont = nFilasCont + 1
        End If
    Next
    
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
End Sub

