Attribute VB_Name = "Módulo1"
Sub Acumular_Meses_1()
    Dim wsResultado As Excel.Worksheet
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim nFilasResultado As Long
    Dim i As Long
    Dim filaCopia As Long
    Dim columnaCopia As Long
    
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Resultados"
    Application.DisplayAlerts = True
    
    Set wsResultado = Worksheets("Resultados")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Fila del encabezado Resultado
    nFilasResultado = 1
    wsResultado.Cells(nFilasResultado, 1).Value = "JurId"
    wsResultado.Cells(nFilasResultado, 2).Value = "Esc"
    wsResultado.Cells(nFilasResultado, 3).Value = "PtaTipo"
    wsResultado.Cells(nFilasResultado, 4).Value = "Doc"
    wsResultado.Cells(nFilasResultado, 5).Value = "Nombres"
    wsResultado.Cells(nFilasResultado, 6).Value = "Año"
    wsResultado.Cells(nFilasResultado, 7).Value = "Enero"
    wsResultado.Cells(nFilasResultado, 8).Value = "Febrero"
    wsResultado.Cells(nFilasResultado, 9).Value = "Marzo"
    wsResultado.Cells(nFilasResultado, 10).Value = "Abril"
    wsResultado.Cells(nFilasResultado, 11).Value = "Mayo"
    wsResultado.Cells(nFilasResultado, 12).Value = "Junio"
    wsResultado.Cells(nFilasResultado, 13).Value = "Julio"
    wsResultado.Cells(nFilasResultado, 14).Value = "SAC cobrado"
    nFilasResultado = 2
    
    repeticiones = 0
    
    For i = 2 To nFilas
        valorJur = Cells(i, 8).Value
        valorDoc = Cells(i, 12).Value
        valorVto = Cells(i, 16).Value
        
        tempMes = Month(valorVto)
        tempAnio = Year(valorVto)
        
        If tempAnio > 2017 Then
            If Cells(i, 4).Value < 300 Or Cells(i, 4).Value = 316 Or Cells(i, 4).Value = 324 Then
                'Busca en la otra hoja
                rangoTemp = "D2:D" & nFilasResultado
                Set resultado = wsResultado.Range(rangoTemp).Find(What:=valorDoc, _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False, _
                            SearchFormat:=False)
                'Si el resultado de la búsqueda no es vacío
                If Not resultado Is Nothing Then
                    'Se obtiene el valor de la fila
                    celdaDoc = resultado.Address
                    tempDoc = ""
                    For m = 1 To Len(celdaDoc)
                        If IsNumeric(Mid(celdaDoc, m, 1)) Then
                            tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                        End If
                    Next m
                    filaCopia = tempDoc
                    
                    If Cells(i, 4).Value = 316 Then
                        If tempMes = 6 Then
                            If Cells(i, 6).Value <> 2 Then
                                wsResultado.Cells(filaCopia, 14).Value = wsResultado.Cells(filaCopia, 14).Value + Cells(i, 7).Value
                            Else
                                wsResultado.Cells(filaCopia, 14).Value = wsResultado.Cells(filaCopia, 14).Value - Cells(i, 7).Value
                            End If
                        End If
                    Else
                        If Cells(i, 6).Value <> 0 Then
                            'Copio en el correspondiente al vto
                            columnaCopia = 6 + tempMes
                            If Cells(i, 6).Value <> 2 Then
                                wsResultado.Cells(filaCopia, columnaCopia).Value = wsResultado.Cells(filaCopia, columnaCopia).Value + Cells(i, 7).Value
                            Else
                                wsResultado.Cells(filaCopia, columnaCopia).Value = wsResultado.Cells(filaCopia, columnaCopia).Value - Cells(i, 7).Value
                            End If
                        Else
                            'Copio en el mes actual
                            columnaCopia = 6 + Cells(i, 2).Value
                            wsResultado.Cells(filaCopia, columnaCopia).Value = wsResultado.Cells(filaCopia, columnaCopia).Value + Cells(i, 7).Value
                        End If
                    End If
                Else
                    'No se encontró el dni. Creo nueva fila
                    wsResultado.Cells(nFilasResultado, 1).Value = Cells(i, 8).Value
                    wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 9).Value
                    wsResultado.Cells(nFilasResultado, 3).Value = Cells(i, 23).Value
                    wsResultado.Cells(nFilasResultado, 4).Value = Cells(i, 12).Value
                    wsResultado.Cells(nFilasResultado, 5).Value = Cells(i, 14).Value
                    wsResultado.Cells(nFilasResultado, 6).Value = Cells(i, 1).Value
                    wsResultado.Cells(nFilasResultado, 7).Value = 0
                    wsResultado.Cells(nFilasResultado, 8).Value = 0
                    wsResultado.Cells(nFilasResultado, 9).Value = 0
                    wsResultado.Cells(nFilasResultado, 10).Value = 0
                    wsResultado.Cells(nFilasResultado, 11).Value = 0
                    wsResultado.Cells(nFilasResultado, 12).Value = 0
                    wsResultado.Cells(nFilasResultado, 13).Value = 0
                    wsResultado.Cells(nFilasResultado, 14).Value = 0
                    
                    If Cells(i, 4).Value = 316 Then
                        If tempMes = 6 Then
                            If Cells(i, 6).Value <> 2 Then
                                wsResultado.Cells(nFilasResultado, 14).Value = wsResultado.Cells(nFilasResultado, 14).Value + Cells(i, 7).Value
                            Else
                                wsResultado.Cells(nFilasResultado, 14).Value = wsResultado.Cells(nFilasResultado, 14).Value - Cells(i, 7).Value
                            End If
                        End If
                    Else
                        If Cells(i, 6).Value <> 0 Then
                            'Copio en el correspondiente al vto
                            columnaCopia = 6 + tempMes
                            If Cells(i, 6).Value <> 2 Then
                                wsResultado.Cells(nFilasResultado, columnaCopia).Value = wsResultado.Cells(nFilasResultado, columnaCopia).Value + Cells(i, 7).Value
                            Else
                                wsResultado.Cells(nFilasResultado, columnaCopia).Value = wsResultado.Cells(nFilasResultado, columnaCopia).Value - Cells(i, 7).Value
                            End If
                        Else
                            'Copio en el mes actual
                            columnaCopia = 6 + Cells(i, 2).Value
                            wsResultado.Cells(nFilasResultado, columnaCopia).Value = wsResultado.Cells(nFilasResultado, columnaCopia).Value + Cells(i, 7).Value
                        End If
                    End If
                    
                    nFilasResultado = nFilasResultado + 1
                End If
            End If
        End If
    Next i
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Acumular_Meses_2()
    Dim wsResultado As Excel.Worksheet
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim nFilasResultado As Long
    Dim i As Long
    Dim filaCopia As Long
    Dim columnaCopia As Long
    
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Resultados"
    Application.DisplayAlerts = True
    
    Set wsResultado = Worksheets("Resultados")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Fila del encabezado Resultado
    nFilasResultado = 1
    wsResultado.Cells(nFilasResultado, 1).Value = "JurId"
    wsResultado.Cells(nFilasResultado, 2).Value = "Esc"
    wsResultado.Cells(nFilasResultado, 3).Value = "PtaTipo"
    wsResultado.Cells(nFilasResultado, 4).Value = "Doc"
    wsResultado.Cells(nFilasResultado, 5).Value = "Nombres"
    wsResultado.Cells(nFilasResultado, 6).Value = "Año"
    wsResultado.Cells(nFilasResultado, 7).Value = "Julio"
    wsResultado.Cells(nFilasResultado, 8).Value = "Agosto"
    wsResultado.Cells(nFilasResultado, 9).Value = "Septiembre"
    wsResultado.Cells(nFilasResultado, 10).Value = "Octubre"
    wsResultado.Cells(nFilasResultado, 11).Value = "Noviembre"
    wsResultado.Cells(nFilasResultado, 12).Value = "Diciembre"
    wsResultado.Cells(nFilasResultado, 13).Value = "SAC cobrado"
    nFilasResultado = 2
    
    
    For i = 2 To nFilas
        valorDoc = Cells(i, 12).Value
        valorVto = Cells(i, 16).Value
        
        tempMes = Month(valorVto)
        tempAnio = Year(valorVto)
        
        If tempAnio > 2016 Then
            If Cells(i, 4).Value < 300 Or Cells(i, 4).Value = 316 Then
                'Busca en la otra hoja
                rangoTemp = "D2:D" & nFilasResultado
                Set resultado = wsResultado.Range(rangoTemp).Find(What:=valorDoc, _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False, _
                            SearchFormat:=False)
                'Si el resultado de la búsqueda no es vacío
                If Not resultado Is Nothing Then
                    'Se obtiene el valor de la fila
                    celdaDoc = resultado.Address
                    tempDoc = ""
                    For m = 1 To Len(celdaDoc)
                        If IsNumeric(Mid(celdaDoc, m, 1)) Then
                            tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                        End If
                    Next m
                    filaCopia = tempDoc
                    
                    If Cells(i, 4).Value = 316 Then
                        If tempMes = 12 Then
                            If Cells(i, 6).Value <> 2 Then
                                wsResultado.Cells(filaCopia, 13).Value = wsResultado.Cells(filaCopia, 13).Value + Cells(i, 7).Value
                            Else
                                wsResultado.Cells(filaCopia, 13).Value = wsResultado.Cells(filaCopia, 13).Value - Cells(i, 7).Value
                            End If
                        End If
                    Else
                        If Cells(i, 6).Value <> 0 Then
                            If tempAnio >= Cells(i, 1).Value And tempMes > Cells(i, 2).Value Then
                                If Cells(i, 2).Value > 6 Then
                                    columnaCopia = Cells(i, 2).Value
                                Else
                                    columnaCopia = 15
                                End If
                            Else
                                If tempMes > 6 Then
                                    columnaCopia = tempMes
                                Else
                                    columnaCopia = 15
                                End If
                            End If
                        Else
                            If Cells(i, 2).Value > 6 Then
                                columnaCopia = Cells(i, 2).Value
                            Else
                                columnaCopia = 15
                            End If
                        End If
                        If Cells(i, 6).Value <> 2 Then
                            wsResultado.Cells(filaCopia, columnaCopia).Value = wsResultado.Cells(filaCopia, columnaCopia).Value + Cells(i, 7).Value
                        Else
                            If tempAnio = 2017 Then
                                wsResultado.Cells(filaCopia, columnaCopia).Value = wsResultado.Cells(filaCopia, columnaCopia).Value - Cells(i, 7).Value
                            End If
                        End If
                    End If
                Else
                    'No se encontró el dni. Creo nueva fila
                    wsResultado.Cells(nFilasResultado, 1).Value = Cells(i, 8).Value
                    wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 9).Value
                    wsResultado.Cells(nFilasResultado, 3).Value = Cells(i, 23).Value
                    wsResultado.Cells(nFilasResultado, 4).Value = Cells(i, 12).Value
                    wsResultado.Cells(nFilasResultado, 5).Value = Cells(i, 14).Value
                    wsResultado.Cells(nFilasResultado, 6).Value = Cells(i, 1).Value
                    wsResultado.Cells(nFilasResultado, 7).Value = 0
                    wsResultado.Cells(nFilasResultado, 8).Value = 0
                    wsResultado.Cells(nFilasResultado, 9).Value = 0
                    wsResultado.Cells(nFilasResultado, 10).Value = 0
                    wsResultado.Cells(nFilasResultado, 11).Value = 0
                    wsResultado.Cells(nFilasResultado, 12).Value = 0
                    wsResultado.Cells(nFilasResultado, 13).Value = 0
                    
                    If Cells(i, 4).Value = 316 Then
                        If tempMes = 12 Then
                            wsResultado.Cells(nFilasResultado, 13).Value = wsResultado.Cells(nFilasResultado, 13).Value + Cells(i, 7).Value
                        End If
                    Else
                        If Cells(i, 6).Value <> 0 Then
                            If tempAnio >= Cells(i, 1).Value And tempMes > Cells(i, 2).Value Then
                                If Cells(i, 2).Value > 6 Then
                                    columnaCopia = Cells(i, 2).Value
                                Else
                                    columnaCopia = 15
                                End If
                            Else
                                If tempMes > 6 Then
                                    columnaCopia = tempMes
                                Else
                                    columnaCopia = 15
                                End If
                            End If
                        Else
                            If Cells(i, 2).Value > 6 Then
                                columnaCopia = Cells(i, 2).Value
                            Else
                                columnaCopia = 15
                            End If
                        End If
                        If Cells(i, 6).Value <> 2 Then
                            wsResultado.Cells(nFilasResultado, columnaCopia).Value = wsResultado.Cells(nFilasResultado, columnaCopia).Value + Cells(i, 7).Value
                        Else
                            If tempAnio = 2017 Then
                                wsResultado.Cells(nFilasResultado, columnaCopia).Value = wsResultado.Cells(nFilasResultado, columnaCopia).Value - Cells(i, 7).Value
                            End If
                        End If
                    End If
                    
                    nFilasResultado = nFilasResultado + 1
                End If
            End If
        End If
    Next i
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub


Sub Generar_Sac()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i As Long
    Dim bandera As Boolean
    Dim importe As Double
    Dim importeSuma As Double
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    For i = 2 To nFilas
        bandera = False
        importe = 0
        importeSuma = 0
        temp = 0
        For j = 7 To 12
            If Cells(i, j).Value = 0 Then
                bandera = True
            Else
                temp = temp + 1
            End If
            If Cells(i, j).Value > importe Then
                importe = Cells(i, j).Value
            End If
            importeSuma = importeSuma + Cells(i, j).Value
        Next j
        If bandera Then
            'Promedio 12
            Cells(i, nColumnas + 1).Value = (importeSuma / 12)
        Else
            'División del mes mayor monto
            Cells(i, nColumnas + 1).Value = importe / 2
        End If
    Next i
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Acumular_Sac()
    Dim wsResultado As Excel.Worksheet
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim nFilasResultado As Long
    Dim i As Long
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Resultados"
    Application.DisplayAlerts = True
    
    Set wsResultado = Worksheets("Resultados")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Fila del encabezado Resultado
    nFilasResultado = 1
    wsResultado.Cells(nFilasResultado, 1).Value = "JurId"
    wsResultado.Cells(nFilasResultado, 2).Value = "Doc"
    wsResultado.Cells(nFilasResultado, 3).Value = "Nombres"
    wsResultado.Cells(nFilasResultado, 4).Value = "Año"
    wsResultado.Cells(nFilasResultado, 5).Value = "Cpto"
    wsResultado.Cells(nFilasResultado, 6).Value = "Importe"
    nFilasResultado = 2
    
    wsResultado.Cells(nFilasResultado, 1).Value = Cells(2, 12).Value
    wsResultado.Cells(nFilasResultado, 2).Value = Cells(2, 5).Value
    wsResultado.Cells(nFilasResultado, 3).Value = Cells(2, 6).Value
    wsResultado.Cells(nFilasResultado, 4).Value = Cells(2, 1).Value
    wsResultado.Cells(nFilasResultado, 5).Value = Cells(2, 7).Value
    wsResultado.Cells(nFilasResultado, 6).Value = Cells(2, 11).Value
    
    For i = 3 To nFilas
        If wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 5).Value Then
            wsResultado.Cells(nFilasResultado, 6).Value = wsResultado.Cells(nFilasResultado, 6).Value + Cells(i, 11).Value
        Else
            nFilasResultado = nFilasResultado + 1
            wsResultado.Cells(nFilasResultado, 1).Value = Cells(i, 12).Value
            wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 5).Value
            wsResultado.Cells(nFilasResultado, 3).Value = Cells(i, 6).Value
            wsResultado.Cells(nFilasResultado, 4).Value = Cells(i, 1).Value
            wsResultado.Cells(nFilasResultado, 5).Value = Cells(i, 7).Value
            wsResultado.Cells(nFilasResultado, 6).Value = Cells(i, 11).Value
        End If
    Next i
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub


Sub Diferencia_Sac()
    Dim wbContenido As Workbook, _
    wsContenido As Excel.Worksheet
    Dim nFilasCont As Long
    Dim nFilas As Long
    Dim nColumnas As Integer
    
    
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
    
    nColumnas = nColumnas + 1
    Cells(1, nColumnas).Value = "SAC Cobrado"
    Cells(1, nColumnas + 1).Value = "Diferencia"
    
    For i = 2 To nFilas
        'Busca en el otro archivo
        valorDoc = Cells(i, 2).Value
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
            'Se obtiene el valor de j
            celdaDoc = resultado.Address
            tempDoc = ""
            For m = 1 To Len(celdaDoc)
                If IsNumeric(Mid(celdaDoc, m, 1)) Then
                    tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                End If
            Next m
            j = tempDoc
            Cells(i, nColumnas).Value = wsContenido.Cells(j, 6).Value
        End If
    Next i
    
    nColumnas = nColumnas + 1
    For i = 2 To nFilas
        If Cells(i, nColumnas - 1).Value <> "" Then
            Cells(i, nColumnas).Value = Cells(i, 11).Value - Cells(i, nColumnas - 1).Value
        Else
            Cells(i, nColumnas).Value = Cells(i, 11).Value
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


Sub Control_Diferencia()
    Dim valorDoc As String
    Dim nFilas As Long
    Dim nColumnas As Integer
    Dim i As Long
    Dim bandera As Boolean
    Dim monto As Double
    
    'Activar este libro
    ThisWorkbook.Activate

    'Regresa el control a la hoja de origen
    Sheets(1).Select

    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    Cells(1, nColumnas + 1).Value = "Observación"
    Cells(1, nColumnas + 2).Value = "Diferencia Mayor"
            
    For i = 2 To nFilas
        bandera = False
        monto = 0
        For j = 13 To nColumnas
            If Cells(i, j - 1).Value > Cells(i, j).Value And Cells(i, j).Value > 0 Then
                If (Cells(i, j - 1).Value - Cells(i, j).Value) > monto And (Cells(i, j - 1).Value - Cells(i, j).Value) > 5 Then
                    monto = Cells(i, j - 1).Value - Cells(i, j).Value
                    bandera = True
                End If
            End If
        Next j
        If bandera Then
            Cells(i, nColumnas + 1).Value = "Controlar"
            Cells(i, nColumnas + 2).Value = monto
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub
