Attribute VB_Name = "Módulo1"
Sub Control_Antiguedad()
    Dim wsResultado As Excel.Worksheet
    Dim rango As Range
    Dim nFilas As Long
    Dim nFilasCont As Long
    Dim valorDoc As String
    Dim unidad As Double
    
    MsgBox "Debe estar ordenado por DNI y Año.", , "¡Atención!"
    
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
    
    nFilasCont = 1
    wsResultado.Cells(nFilasCont, 1).Value = "JUR"
    wsResultado.Cells(nFilasCont, 2).Value = "DNI"
    wsResultado.Cells(nFilasCont, 3).Value = "NOMBRE"
    wsResultado.Cells(nFilasCont, 4).Value = "CEIC"
    wsResultado.Cells(nFilasCont, 5).Value = "AÑO CONTROL"
    wsResultado.Cells(nFilasCont, 6).Value = "DIFERENCIA"
    nFilasCont = 2
    

    For i = 3 To (nFilas - 1)
        valorDoc = Cells(i, 12).Value
        unidad = Cells(i, 18).Value
        j = i + 1
        If valorDoc = Cells(j, 12).Value Then
            If Cells(j, 18).Value - unidad <> 1 Then
                wsResultado.Cells(nFilasCont, 1).Value = Cells(i, 8).Value
                wsResultado.Cells(nFilasCont, 2).Value = Cells(i, 12).Value
                wsResultado.Cells(nFilasCont, 3).Value = Cells(i, 14).Value
                wsResultado.Cells(nFilasCont, 4).Value = Cells(i, 15).Value
                wsResultado.Cells(nFilasCont, 5).Value = Cells(i, 1).Value & " - " & Cells(j, 1).Value
                wsResultado.Cells(nFilasCont, 6).Value = Cells(j, 18).Value - unidad
                nFilasCont = nFilasCont + 1
            End If
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub


Sub Informe_Antiguedad()
    Dim wsResultado As Excel.Worksheet
    Dim rango As Range
    Dim nFilas As Long
    Dim nFilasCont As Long
    Dim valorDoc As String
    Dim jur As Integer
    Dim cantidad As Long
    Dim cobran As Long
    
    MsgBox "Debe estar ordenado por JUR y DNI.", , "¡Atención!"
    
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
    
    nFilasCont = 1
    wsResultado.Cells(nFilasCont, 1).Value = "JUR"
    wsResultado.Cells(nFilasCont, 2).Value = "Rango"
    wsResultado.Cells(nFilasCont, 3).Value = "Total"
    wsResultado.Cells(nFilasCont, 4).Value = "Porcentaje"
    
    jur = Cells(2, 8).Value
    cantidad = 0
    cobran = 0
    For j = 1 To 7
        wsResultado.Cells(nFilasCont + j, 1).Value = jur
        wsResultado.Cells(nFilasCont + j, 3).Value = 0
    Next j
    wsResultado.Cells(nFilasCont + 1, 2).Value = ".1 - 9"
    wsResultado.Cells(nFilasCont + 2, 2).Value = ".10 - 14"
    wsResultado.Cells(nFilasCont + 3, 2).Value = ".15 - 19"
    wsResultado.Cells(nFilasCont + 4, 2).Value = ".20 - 24"
    wsResultado.Cells(nFilasCont + 5, 2).Value = ".25 - 29"
    wsResultado.Cells(nFilasCont + 6, 2).Value = ".30 o más"
    wsResultado.Cells(nFilasCont + 7, 2).Value = "no cobran"
    
    For i = 2 To nFilas
        If jur = Cells(i, 8).Value Then
            If Cells(i, 4).Value = 200 Then
                cobran = cobran + 1
                If Cells(i, 18).Value >= 30 Then
                    wsResultado.Cells(nFilasCont + 6, 3).Value = wsResultado.Cells(nFilasCont + 6, 3).Value + 1
                Else
                    If Cells(i, 18).Value >= 25 Then
                        wsResultado.Cells(nFilasCont + 5, 3).Value = wsResultado.Cells(nFilasCont + 5, 3).Value + 1
                    Else
                        If Cells(i, 18).Value >= 20 Then
                            wsResultado.Cells(nFilasCont + 4, 3).Value = wsResultado.Cells(nFilasCont + 4, 3).Value + 1
                        Else
                            If Cells(i, 18).Value >= 15 Then
                                wsResultado.Cells(nFilasCont + 3, 3).Value = wsResultado.Cells(nFilasCont + 3, 3).Value + 1
                            Else
                                If Cells(i, 18).Value >= 10 Then
                                    wsResultado.Cells(nFilasCont + 2, 3).Value = wsResultado.Cells(nFilasCont + 2, 3).Value + 1
                                Else
                                    wsResultado.Cells(nFilasCont + 1, 3).Value = wsResultado.Cells(nFilasCont + 1, 3).Value + 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                cantidad = cantidad + 1
            End If
        Else
            wsResultado.Cells(nFilasCont + 7, 3).Value = cantidad - cobran
            nFilasCont = nFilasCont + 8
            wsResultado.Cells(nFilasCont, 2).Value = "TOTAL"
            temp = "B" & nFilasCont
            wsResultado.Range(temp).Font.Bold = True
            wsResultado.Cells(nFilasCont, 3).Value = cantidad
            'Calcular porcentajes
            wsResultado.Cells(nFilasCont, 4).Value = 1
            wsResultado.Cells(nFilasCont - 1, 4).Value = wsResultado.Cells(nFilasCont - 1, 3).Value / cantidad
            wsResultado.Cells(nFilasCont - 2, 4).Value = wsResultado.Cells(nFilasCont - 2, 3).Value / cantidad
            wsResultado.Cells(nFilasCont - 3, 4).Value = wsResultado.Cells(nFilasCont - 3, 3).Value / cantidad
            wsResultado.Cells(nFilasCont - 4, 4).Value = wsResultado.Cells(nFilasCont - 4, 3).Value / cantidad
            wsResultado.Cells(nFilasCont - 5, 4).Value = wsResultado.Cells(nFilasCont - 5, 3).Value / cantidad
            wsResultado.Cells(nFilasCont - 6, 4).Value = wsResultado.Cells(nFilasCont - 6, 3).Value / cantidad
            wsResultado.Cells(nFilasCont - 7, 4).Value = wsResultado.Cells(nFilasCont - 7, 3).Value / cantidad
            
            jur = Cells(i, 8).Value
            cantidad = 1
            cobran = 0
            For j = 1 To 7
                wsResultado.Cells(nFilasCont + j, 1).Value = jur
                wsResultado.Cells(nFilasCont + j, 3).Value = 0
            Next j
            wsResultado.Cells(nFilasCont + 1, 2).Value = "1 - 9"
            wsResultado.Cells(nFilasCont + 2, 2).Value = "10 - 14"
            wsResultado.Cells(nFilasCont + 3, 2).Value = "15 - 19"
            wsResultado.Cells(nFilasCont + 4, 2).Value = "20 - 24"
            wsResultado.Cells(nFilasCont + 5, 2).Value = "25 - 29"
            wsResultado.Cells(nFilasCont + 6, 2).Value = "30 o más"
            wsResultado.Cells(nFilasCont + 7, 2).Value = "no cobran"
            
            'trato el primero
            'valorDoc = Cells(i, 12).Value
            If Cells(i, 4).Value = 200 Then
                cobran = 1
                If Cells(i, 18).Value >= 30 Then
                    wsResultado.Cells(nFilasCont + 6, 3).Value = 1
                Else
                    If Cells(i, 18).Value >= 25 Then
                        wsResultado.Cells(nFilasCont + 5, 3).Value = 1
                    Else
                        If Cells(i, 18).Value >= 20 Then
                            wsResultado.Cells(nFilasCont + 4, 3).Value = 1
                        Else
                            If Cells(i, 18).Value >= 15 Then
                                wsResultado.Cells(nFilasCont + 3, 3).Value = 1
                            Else
                                If Cells(i, 18).Value >= 10 Then
                                    wsResultado.Cells(nFilasCont + 2, 3).Value = 1
                                Else
                                    wsResultado.Cells(nFilasCont + 1, 3).Value = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
    Next i
    
    wsResultado.Cells(nFilasCont + 7, 3).Value = cantidad - cobran
    nFilasCont = nFilasCont + 8
    wsResultado.Cells(nFilasCont, 2).Value = "TOTAL"
    temp = "B" & nFilasCont
    wsResultado.Range(temp).Font.Bold = True
    wsResultado.Cells(nFilasCont, 3).Value = cantidad
    'Calcular porcentajes
    wsResultado.Cells(nFilasCont, 4).Value = 1
    wsResultado.Cells(nFilasCont - 1, 4).Value = wsResultado.Cells(nFilasCont - 1, 3).Value / cantidad
    wsResultado.Cells(nFilasCont - 2, 4).Value = wsResultado.Cells(nFilasCont - 2, 3).Value / cantidad
    wsResultado.Cells(nFilasCont - 3, 4).Value = wsResultado.Cells(nFilasCont - 3, 3).Value / cantidad
    wsResultado.Cells(nFilasCont - 4, 4).Value = wsResultado.Cells(nFilasCont - 4, 3).Value / cantidad
    wsResultado.Cells(nFilasCont - 5, 4).Value = wsResultado.Cells(nFilasCont - 5, 3).Value / cantidad
    wsResultado.Cells(nFilasCont - 6, 4).Value = wsResultado.Cells(nFilasCont - 6, 3).Value / cantidad
    wsResultado.Cells(nFilasCont - 7, 4).Value = wsResultado.Cells(nFilasCont - 7, 3).Value / cantidad
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub


Sub Informe_Antiguedad_Edad()
    Dim wsResultado As Excel.Worksheet
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim nFilasCont As Long
    Dim valorDoc As String
    Dim jur As Integer
    Dim cantidad As Long
    Dim cobran As Long
    
    MsgBox "Debe estar ordenado por JUR y DNI.", , "¡Atención!"
    
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
    nColumnas = rango.Columns.Count
    
    nFilasCont = 1
    wsResultado.Cells(nFilasCont, 1).Value = "JUR: " & Cells(2, 8).Value
    wsResultado.Range("A1").Font.Bold = True
    
    jur = Cells(2, 8).Value
    cantidad = 0
    cobran = 0
    For j = 2 To 9
        For m = 2 To 12 Step 2
        wsResultado.Cells(nFilasCont + j, m).Value = 0
        Next m
    Next j
    wsResultado.Cells(nFilasCont + 1, 1).Value = "ANTIG\EDAD"
    wsResultado.Cells(nFilasCont + 2, 1).Value = ".1 - 9"
    wsResultado.Cells(nFilasCont + 3, 1).Value = ".10 - 14"
    wsResultado.Cells(nFilasCont + 4, 1).Value = ".15 - 19"
    wsResultado.Cells(nFilasCont + 5, 1).Value = ".20 - 24"
    wsResultado.Cells(nFilasCont + 6, 1).Value = ".25 - 29"
    wsResultado.Cells(nFilasCont + 7, 1).Value = ".30 o más"
    wsResultado.Cells(nFilasCont + 8, 1).Value = "No cobran"
    wsResultado.Cells(nFilasCont + 9, 1).Value = "TOTAL"
    wsResultado.Cells(nFilasCont + 1, 2).Value = "Menor 40"
    wsResultado.Cells(nFilasCont + 1, 4).Value = "40 - 49"
    wsResultado.Cells(nFilasCont + 1, 6).Value = "50 - 59"
    wsResultado.Cells(nFilasCont + 1, 8).Value = "60 - 65"
    wsResultado.Cells(nFilasCont + 1, 10).Value = "Mayor 65"
    wsResultado.Cells(nFilasCont + 1, 12).Value = "TOTAL"
    
    Application.DisplayAlerts = False
    fila = nFilasCont + 1
    wsResultado.Range("B" & fila & ":C" & fila).Merge
    wsResultado.Range("B" & fila & ":C" & fila).HorizontalAlignment = xlCenter
    wsResultado.Range("B" & fila & ":C" & fila).Font.Bold = True
    wsResultado.Range("D" & fila & ":E" & fila).Merge
    wsResultado.Range("D" & fila & ":E" & fila).HorizontalAlignment = xlCenter
    wsResultado.Range("D" & fila & ":E" & fila).Font.Bold = True
    wsResultado.Range("F" & fila & ":G" & fila).Merge
    wsResultado.Range("F" & fila & ":G" & fila).HorizontalAlignment = xlCenter
    wsResultado.Range("F" & fila & ":G" & fila).Font.Bold = True
    wsResultado.Range("H" & fila & ":I" & fila).Merge
    wsResultado.Range("H" & fila & ":I" & fila).HorizontalAlignment = xlCenter
    wsResultado.Range("H" & fila & ":I" & fila).Font.Bold = True
    wsResultado.Range("J" & fila & ":K" & fila).Merge
    wsResultado.Range("J" & fila & ":K" & fila).HorizontalAlignment = xlCenter
    wsResultado.Range("J" & fila & ":K" & fila).Font.Bold = True
    wsResultado.Range("L" & fila & ":M" & fila).Merge
    wsResultado.Range("L" & fila & ":M" & fila).HorizontalAlignment = xlCenter
    wsResultado.Range("L" & fila & ":M" & fila).Font.Bold = True
    Application.DisplayAlerts = True
    
    For i = 2 To nFilas
        If jur = Cells(i, 8).Value Then
            If Cells(i, 4).Value = 200 Then
                cobran = cobran + 1
                If Cells(i, 18).Value >= 30 Then
                    If Cells(i, nColumnas).Value < 40 Then
                        wsResultado.Cells(nFilasCont + 7, 2).Value = wsResultado.Cells(nFilasCont + 7, 2).Value + 1
                        wsResultado.Cells(nFilasCont + 9, 2).Value = wsResultado.Cells(nFilasCont + 9, 2).Value + 1
                    Else
                        If Cells(i, nColumnas).Value < 50 Then
                            wsResultado.Cells(nFilasCont + 7, 4).Value = wsResultado.Cells(nFilasCont + 7, 4).Value + 1
                            wsResultado.Cells(nFilasCont + 9, 4).Value = wsResultado.Cells(nFilasCont + 9, 4).Value + 1
                        Else
                            If Cells(i, nColumnas).Value < 60 Then
                                wsResultado.Cells(nFilasCont + 7, 6).Value = wsResultado.Cells(nFilasCont + 7, 6).Value + 1
                                wsResultado.Cells(nFilasCont + 9, 6).Value = wsResultado.Cells(nFilasCont + 9, 6).Value + 1
                            Else
                                If Cells(i, nColumnas).Value < 66 Then
                                    wsResultado.Cells(nFilasCont + 7, 8).Value = wsResultado.Cells(nFilasCont + 7, 8).Value + 1
                                    wsResultado.Cells(nFilasCont + 9, 8).Value = wsResultado.Cells(nFilasCont + 9, 8).Value + 1
                                Else
                                    wsResultado.Cells(nFilasCont + 7, 10).Value = wsResultado.Cells(nFilasCont + 7, 10).Value + 1
                                    wsResultado.Cells(nFilasCont + 9, 10).Value = wsResultado.Cells(nFilasCont + 9, 10).Value + 1
                                End If
                            End If
                        End If
                    End If
                    wsResultado.Cells(nFilasCont + 7, 12).Value = wsResultado.Cells(nFilasCont + 7, 12).Value + 1
                Else
                    If Cells(i, 18).Value >= 25 Then
                        If Cells(i, nColumnas).Value < 40 Then
                            wsResultado.Cells(nFilasCont + 6, 2).Value = wsResultado.Cells(nFilasCont + 6, 2).Value + 1
                            wsResultado.Cells(nFilasCont + 9, 2).Value = wsResultado.Cells(nFilasCont + 9, 2).Value + 1
                        Else
                            If Cells(i, nColumnas).Value < 50 Then
                                wsResultado.Cells(nFilasCont + 6, 4).Value = wsResultado.Cells(nFilasCont + 6, 4).Value + 1
                                wsResultado.Cells(nFilasCont + 9, 4).Value = wsResultado.Cells(nFilasCont + 9, 4).Value + 1
                            Else
                                If Cells(i, nColumnas).Value < 60 Then
                                    wsResultado.Cells(nFilasCont + 6, 6).Value = wsResultado.Cells(nFilasCont + 6, 6).Value + 1
                                    wsResultado.Cells(nFilasCont + 9, 6).Value = wsResultado.Cells(nFilasCont + 9, 6).Value + 1
                                Else
                                    If Cells(i, nColumnas).Value < 66 Then
                                        wsResultado.Cells(nFilasCont + 6, 8).Value = wsResultado.Cells(nFilasCont + 6, 8).Value + 1
                                        wsResultado.Cells(nFilasCont + 9, 8).Value = wsResultado.Cells(nFilasCont + 9, 8).Value + 1
                                    Else
                                        wsResultado.Cells(nFilasCont + 6, 10).Value = wsResultado.Cells(nFilasCont + 6, 10).Value + 1
                                        wsResultado.Cells(nFilasCont + 9, 10).Value = wsResultado.Cells(nFilasCont + 9, 10).Value + 1
                                    End If
                                End If
                            End If
                        End If
                        wsResultado.Cells(nFilasCont + 6, 12).Value = wsResultado.Cells(nFilasCont + 6, 12).Value + 1
                    Else
                        If Cells(i, 18).Value >= 20 Then
                            If Cells(i, nColumnas).Value < 40 Then
                                wsResultado.Cells(nFilasCont + 5, 2).Value = wsResultado.Cells(nFilasCont + 5, 2).Value + 1
                                wsResultado.Cells(nFilasCont + 9, 2).Value = wsResultado.Cells(nFilasCont + 9, 2).Value + 1
                            Else
                                If Cells(i, nColumnas).Value < 50 Then
                                    wsResultado.Cells(nFilasCont + 5, 4).Value = wsResultado.Cells(nFilasCont + 5, 4).Value + 1
                                    wsResultado.Cells(nFilasCont + 9, 4).Value = wsResultado.Cells(nFilasCont + 9, 4).Value + 1
                                Else
                                    If Cells(i, nColumnas).Value < 60 Then
                                        wsResultado.Cells(nFilasCont + 5, 6).Value = wsResultado.Cells(nFilasCont + 5, 6).Value + 1
                                        wsResultado.Cells(nFilasCont + 9, 6).Value = wsResultado.Cells(nFilasCont + 9, 6).Value + 1
                                    Else
                                        If Cells(i, nColumnas).Value < 66 Then
                                            wsResultado.Cells(nFilasCont + 5, 8).Value = wsResultado.Cells(nFilasCont + 5, 8).Value + 1
                                            wsResultado.Cells(nFilasCont + 9, 8).Value = wsResultado.Cells(nFilasCont + 9, 8).Value + 1
                                        Else
                                            wsResultado.Cells(nFilasCont + 5, 10).Value = wsResultado.Cells(nFilasCont + 5, 10).Value + 1
                                            wsResultado.Cells(nFilasCont + 9, 10).Value = wsResultado.Cells(nFilasCont + 9, 10).Value + 1
                                        End If
                                    End If
                                End If
                            End If
                            wsResultado.Cells(nFilasCont + 5, 12).Value = wsResultado.Cells(nFilasCont + 5, 12).Value + 1
                        Else
                            If Cells(i, 18).Value >= 15 Then
                                If Cells(i, nColumnas).Value < 40 Then
                                    wsResultado.Cells(nFilasCont + 4, 2).Value = wsResultado.Cells(nFilasCont + 4, 2).Value + 1
                                    wsResultado.Cells(nFilasCont + 9, 2).Value = wsResultado.Cells(nFilasCont + 9, 2).Value + 1
                                Else
                                    If Cells(i, nColumnas).Value < 50 Then
                                        wsResultado.Cells(nFilasCont + 4, 4).Value = wsResultado.Cells(nFilasCont + 4, 4).Value + 1
                                        wsResultado.Cells(nFilasCont + 9, 4).Value = wsResultado.Cells(nFilasCont + 9, 4).Value + 1
                                    Else
                                        If Cells(i, nColumnas).Value < 60 Then
                                            wsResultado.Cells(nFilasCont + 4, 6).Value = wsResultado.Cells(nFilasCont + 4, 6).Value + 1
                                            wsResultado.Cells(nFilasCont + 9, 6).Value = wsResultado.Cells(nFilasCont + 9, 6).Value + 1
                                        Else
                                            If Cells(i, nColumnas).Value < 66 Then
                                                wsResultado.Cells(nFilasCont + 4, 8).Value = wsResultado.Cells(nFilasCont + 4, 8).Value + 1
                                                wsResultado.Cells(nFilasCont + 9, 8).Value = wsResultado.Cells(nFilasCont + 9, 8).Value + 1
                                            Else
                                                wsResultado.Cells(nFilasCont + 4, 10).Value = wsResultado.Cells(nFilasCont + 4, 10).Value + 1
                                                wsResultado.Cells(nFilasCont + 9, 10).Value = wsResultado.Cells(nFilasCont + 9, 10).Value + 1
                                            End If
                                        End If
                                    End If
                                End If
                                wsResultado.Cells(nFilasCont + 4, 12).Value = wsResultado.Cells(nFilasCont + 4, 12).Value + 1
                            Else
                                If Cells(i, 18).Value >= 10 Then
                                    If Cells(i, nColumnas).Value < 40 Then
                                        wsResultado.Cells(nFilasCont + 3, 2).Value = wsResultado.Cells(nFilasCont + 3, 2).Value + 1
                                        wsResultado.Cells(nFilasCont + 9, 2).Value = wsResultado.Cells(nFilasCont + 9, 2).Value + 1
                                    Else
                                        If Cells(i, nColumnas).Value < 50 Then
                                            wsResultado.Cells(nFilasCont + 3, 4).Value = wsResultado.Cells(nFilasCont + 3, 4).Value + 1
                                            wsResultado.Cells(nFilasCont + 9, 4).Value = wsResultado.Cells(nFilasCont + 9, 4).Value + 1
                                        Else
                                            If Cells(i, nColumnas).Value < 60 Then
                                                wsResultado.Cells(nFilasCont + 3, 6).Value = wsResultado.Cells(nFilasCont + 3, 6).Value + 1
                                                wsResultado.Cells(nFilasCont + 9, 6).Value = wsResultado.Cells(nFilasCont + 9, 6).Value + 1
                                            Else
                                                If Cells(i, nColumnas).Value < 66 Then
                                                    wsResultado.Cells(nFilasCont + 3, 8).Value = wsResultado.Cells(nFilasCont + 3, 8).Value + 1
                                                    wsResultado.Cells(nFilasCont + 9, 8).Value = wsResultado.Cells(nFilasCont + 9, 8).Value + 1
                                                Else
                                                    wsResultado.Cells(nFilasCont + 3, 10).Value = wsResultado.Cells(nFilasCont + 3, 10).Value + 1
                                                    wsResultado.Cells(nFilasCont + 9, 10).Value = wsResultado.Cells(nFilasCont + 9, 10).Value + 1
                                                End If
                                            End If
                                        End If
                                    End If
                                    wsResultado.Cells(nFilasCont + 3, 12).Value = wsResultado.Cells(nFilasCont + 3, 12).Value + 1
                                Else
                                    If Cells(i, nColumnas).Value < 40 Then
                                        wsResultado.Cells(nFilasCont + 2, 2).Value = wsResultado.Cells(nFilasCont + 2, 2).Value + 1
                                        wsResultado.Cells(nFilasCont + 9, 2).Value = wsResultado.Cells(nFilasCont + 9, 2).Value + 1
                                    Else
                                        If Cells(i, nColumnas).Value < 50 Then
                                            wsResultado.Cells(nFilasCont + 2, 4).Value = wsResultado.Cells(nFilasCont + 2, 4).Value + 1
                                            wsResultado.Cells(nFilasCont + 9, 4).Value = wsResultado.Cells(nFilasCont + 9, 4).Value + 1
                                        Else
                                            If Cells(i, nColumnas).Value < 60 Then
                                                wsResultado.Cells(nFilasCont + 2, 6).Value = wsResultado.Cells(nFilasCont + 2, 6).Value + 1
                                                wsResultado.Cells(nFilasCont + 9, 6).Value = wsResultado.Cells(nFilasCont + 9, 6).Value + 1
                                            Else
                                                If Cells(i, nColumnas).Value < 66 Then
                                                    wsResultado.Cells(nFilasCont + 2, 8).Value = wsResultado.Cells(nFilasCont + 2, 8).Value + 1
                                                    wsResultado.Cells(nFilasCont + 9, 8).Value = wsResultado.Cells(nFilasCont + 9, 8).Value + 1
                                                Else
                                                    wsResultado.Cells(nFilasCont + 2, 10).Value = wsResultado.Cells(nFilasCont + 2, 10).Value + 1
                                                    wsResultado.Cells(nFilasCont + 9, 10).Value = wsResultado.Cells(nFilasCont + 9, 10).Value + 1
                                                End If
                                            End If
                                        End If
                                    End If
                                    wsResultado.Cells(nFilasCont + 2, 12).Value = wsResultado.Cells(nFilasCont + 2, 12).Value + 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                'es cpto 1
                cantidad = cantidad + 1
                If Cells(i - 1, 12).Value <> Cells(i, 12).Value And Cells(i + 1, 12).Value <> Cells(i, 12).Value Then
                    'el dni no tiene cpto 200
                    If Cells(i, nColumnas).Value < 40 Then
                        wsResultado.Cells(nFilasCont + 8, 2).Value = wsResultado.Cells(nFilasCont + 8, 2).Value + 1
                        wsResultado.Cells(nFilasCont + 9, 2).Value = wsResultado.Cells(nFilasCont + 9, 2).Value + 1
                    Else
                        If Cells(i, nColumnas).Value < 50 Then
                            wsResultado.Cells(nFilasCont + 8, 4).Value = wsResultado.Cells(nFilasCont + 8, 4).Value + 1
                            wsResultado.Cells(nFilasCont + 9, 4).Value = wsResultado.Cells(nFilasCont + 9, 4).Value + 1
                        Else
                            If Cells(i, nColumnas).Value < 60 Then
                                wsResultado.Cells(nFilasCont + 8, 6).Value = wsResultado.Cells(nFilasCont + 8, 6).Value + 1
                                wsResultado.Cells(nFilasCont + 9, 6).Value = wsResultado.Cells(nFilasCont + 9, 6).Value + 1
                            Else
                                If Cells(i, nColumnas).Value < 66 Then
                                    wsResultado.Cells(nFilasCont + 8, 8).Value = wsResultado.Cells(nFilasCont + 8, 8).Value + 1
                                    wsResultado.Cells(nFilasCont + 9, 8).Value = wsResultado.Cells(nFilasCont + 9, 8).Value + 1
                                Else
                                    wsResultado.Cells(nFilasCont + 8, 10).Value = wsResultado.Cells(nFilasCont + 8, 10).Value + 1
                                    wsResultado.Cells(nFilasCont + 9, 10).Value = wsResultado.Cells(nFilasCont + 9, 10).Value + 1
                                End If
                            End If
                        End If
                    End If
                    wsResultado.Cells(nFilasCont + 8, 12).Value = wsResultado.Cells(nFilasCont + 8, 12).Value + 1
                End If
            End If
        Else
            'Calcular %
            wsResultado.Cells(nFilasCont + 9, 12).Value = cantidad
            For m = 2 To 12 Step 2
                For j = 2 To 9
                    wsResultado.Cells(nFilasCont + j, m + 1).Value = wsResultado.Cells(nFilasCont + j, m).Value / cantidad
                Next j
            Next m
            
            nFilasCont = nFilasCont + 11
            jur = Cells(i, 8).Value
            cantidad = 0
            cobran = 0
            
            wsResultado.Cells(nFilasCont, 1).Value = "JUR: " & jur
            temp = "A" & nFilasCont
            wsResultado.Range(temp).Font.Bold = True
    
            For j = 2 To 9
                For m = 2 To 12 Step 2
                wsResultado.Cells(nFilasCont + j, m).Value = 0
                Next m
            Next j
            wsResultado.Cells(nFilasCont + 1, 1).Value = "ANTIG\EDAD"
            wsResultado.Cells(nFilasCont + 2, 1).Value = ".1 - 9"
            wsResultado.Cells(nFilasCont + 3, 1).Value = ".10 - 14"
            wsResultado.Cells(nFilasCont + 4, 1).Value = ".15 - 19"
            wsResultado.Cells(nFilasCont + 5, 1).Value = ".20 - 24"
            wsResultado.Cells(nFilasCont + 6, 1).Value = ".25 - 29"
            wsResultado.Cells(nFilasCont + 7, 1).Value = ".30 o más"
            wsResultado.Cells(nFilasCont + 8, 1).Value = "No cobran"
            wsResultado.Cells(nFilasCont + 9, 1).Value = "TOTAL"
            wsResultado.Cells(nFilasCont + 1, 2).Value = "Menor 40"
            wsResultado.Cells(nFilasCont + 1, 4).Value = "40 - 49"
            wsResultado.Cells(nFilasCont + 1, 6).Value = "50 - 59"
            wsResultado.Cells(nFilasCont + 1, 8).Value = "60 - 65"
            wsResultado.Cells(nFilasCont + 1, 10).Value = "Mayor 65"
            wsResultado.Cells(nFilasCont + 1, 12).Value = "TOTAL"
            
            Application.DisplayAlerts = False
            fila = nFilasCont + 1
            wsResultado.Range("B" & fila & ":C" & fila).Merge
            wsResultado.Range("B" & fila & ":C" & fila).HorizontalAlignment = xlCenter
            wsResultado.Range("B" & fila & ":C" & fila).Font.Bold = True
            wsResultado.Range("D" & fila & ":E" & fila).Merge
            wsResultado.Range("D" & fila & ":E" & fila).HorizontalAlignment = xlCenter
            wsResultado.Range("D" & fila & ":E" & fila).Font.Bold = True
            wsResultado.Range("F" & fila & ":G" & fila).Merge
            wsResultado.Range("F" & fila & ":G" & fila).HorizontalAlignment = xlCenter
            wsResultado.Range("F" & fila & ":G" & fila).Font.Bold = True
            wsResultado.Range("H" & fila & ":I" & fila).Merge
            wsResultado.Range("H" & fila & ":I" & fila).HorizontalAlignment = xlCenter
            wsResultado.Range("H" & fila & ":I" & fila).Font.Bold = True
            wsResultado.Range("J" & fila & ":K" & fila).Merge
            wsResultado.Range("J" & fila & ":K" & fila).HorizontalAlignment = xlCenter
            wsResultado.Range("J" & fila & ":K" & fila).Font.Bold = True
            wsResultado.Range("L" & fila & ":M" & fila).Merge
            wsResultado.Range("L" & fila & ":M" & fila).HorizontalAlignment = xlCenter
            wsResultado.Range("L" & fila & ":M" & fila).Font.Bold = True
            Application.DisplayAlerts = True
            
            i = i - 1
        End If
        
    Next i
    
    'Calcular %
    wsResultado.Cells(nFilasCont + 9, 12).Value = cantidad
    For m = 2 To 12 Step 2
        For j = 2 To 9
            wsResultado.Cells(nFilasCont + j, m + 1).Value = wsResultado.Cells(nFilasCont + j, m).Value / cantidad
        Next j
    Next m
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Informe_Antiguedad_Edad_2()
    Dim wsResultado As Excel.Worksheet
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim nFilasCont As Long
    Dim valorDoc As String
    Dim jur As Integer
    Dim cantidad As Long
    Dim cobran As Long
    
    MsgBox "Debe estar ordenado por JUR y DNI.", , "¡Atención!"
    
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
    nColumnas = rango.Columns.Count
    
    nFilasCont = 1
    wsResultado.Cells(nFilasCont, 1).Value = "JUR: " & Cells(2, 8).Value
    wsResultado.Range("A1").Font.Bold = True
    
    jur = Cells(2, 8).Value
    cantidad = 0
    cobran = 0
    For j = 2 To 9
        For m = 2 To 12 Step 2
        wsResultado.Cells(nFilasCont + j, m).Value = 0
        Next m
    Next j
    wsResultado.Cells(nFilasCont + 1, 1).Value = "ANTIG\EDAD"
    wsResultado.Cells(nFilasCont + 2, 1).Value = ".1 - 9"
    wsResultado.Cells(nFilasCont + 3, 1).Value = ".10 - 14"
    wsResultado.Cells(nFilasCont + 4, 1).Value = ".15 - 19"
    wsResultado.Cells(nFilasCont + 5, 1).Value = ".20 - 24"
    wsResultado.Cells(nFilasCont + 6, 1).Value = ".25 - 29"
    wsResultado.Cells(nFilasCont + 7, 1).Value = ".30 o más"
    wsResultado.Cells(nFilasCont + 8, 1).Value = "No cobran"
    wsResultado.Cells(nFilasCont + 9, 1).Value = "TOTAL"
    wsResultado.Cells(nFilasCont + 1, 2).Value = "Menor 40"
    wsResultado.Cells(nFilasCont + 1, 4).Value = "40 - 44"
    wsResultado.Cells(nFilasCont + 1, 6).Value = "45 - 49"
    wsResultado.Cells(nFilasCont + 1, 8).Value = "50 - 54"
    wsResultado.Cells(nFilasCont + 1, 10).Value = "55 - 59"
    wsResultado.Cells(nFilasCont + 1, 12).Value = "60 - 65"
    wsResultado.Cells(nFilasCont + 1, 14).Value = "Mayor 65"
    wsResultado.Cells(nFilasCont + 1, 16).Value = "TOTAL"
    
    Application.DisplayAlerts = False
    fila = nFilasCont + 1
    wsResultado.Range("B" & fila & ":C" & fila).Merge
    wsResultado.Range("B" & fila & ":C" & fila).HorizontalAlignment = xlCenter
    wsResultado.Range("B" & fila & ":C" & fila).Font.Bold = True
    wsResultado.Range("D" & fila & ":E" & fila).Merge
    wsResultado.Range("D" & fila & ":E" & fila).HorizontalAlignment = xlCenter
    wsResultado.Range("D" & fila & ":E" & fila).Font.Bold = True
    wsResultado.Range("F" & fila & ":G" & fila).Merge
    wsResultado.Range("F" & fila & ":G" & fila).HorizontalAlignment = xlCenter
    wsResultado.Range("F" & fila & ":G" & fila).Font.Bold = True
    wsResultado.Range("H" & fila & ":I" & fila).Merge
    wsResultado.Range("H" & fila & ":I" & fila).HorizontalAlignment = xlCenter
    wsResultado.Range("H" & fila & ":I" & fila).Font.Bold = True
    wsResultado.Range("J" & fila & ":K" & fila).Merge
    wsResultado.Range("J" & fila & ":K" & fila).HorizontalAlignment = xlCenter
    wsResultado.Range("J" & fila & ":K" & fila).Font.Bold = True
    wsResultado.Range("L" & fila & ":M" & fila).Merge
    wsResultado.Range("L" & fila & ":M" & fila).HorizontalAlignment = xlCenter
    wsResultado.Range("L" & fila & ":M" & fila).Font.Bold = True
    wsResultado.Range("N" & fila & ":O" & fila).Merge
    wsResultado.Range("N" & fila & ":O" & fila).HorizontalAlignment = xlCenter
    wsResultado.Range("N" & fila & ":O" & fila).Font.Bold = True
    wsResultado.Range("P" & fila & ":Q" & fila).Merge
    wsResultado.Range("P" & fila & ":Q" & fila).HorizontalAlignment = xlCenter
    wsResultado.Range("P" & fila & ":Q" & fila).Font.Bold = True
    wsResultado.Range("R" & fila & ":S" & fila).Merge
    wsResultado.Range("R" & fila & ":S" & fila).HorizontalAlignment = xlCenter
    wsResultado.Range("R" & fila & ":S" & fila).Font.Bold = True
    Application.DisplayAlerts = True
    
    For i = 2 To nFilas
        If jur = Cells(i, 8).Value Then
            If Cells(i, 4).Value = 200 Then
                cobran = cobran + 1
                If Cells(i, 18).Value >= 30 Then
                    If Cells(i, nColumnas).Value < 40 Then
                        wsResultado.Cells(nFilasCont + 7, 2).Value = wsResultado.Cells(nFilasCont + 7, 2).Value + 1
                        wsResultado.Cells(nFilasCont + 9, 2).Value = wsResultado.Cells(nFilasCont + 9, 2).Value + 1
                    Else
                        If Cells(i, nColumnas).Value < 45 Then
                            wsResultado.Cells(nFilasCont + 7, 4).Value = wsResultado.Cells(nFilasCont + 7, 4).Value + 1
                            wsResultado.Cells(nFilasCont + 9, 4).Value = wsResultado.Cells(nFilasCont + 9, 4).Value + 1
                        Else
                            If Cells(i, nColumnas).Value < 50 Then
                                wsResultado.Cells(nFilasCont + 7, 6).Value = wsResultado.Cells(nFilasCont + 7, 6).Value + 1
                                wsResultado.Cells(nFilasCont + 9, 6).Value = wsResultado.Cells(nFilasCont + 9, 6).Value + 1
                            Else
                                If Cells(i, nColumnas).Value < 55 Then
                                    wsResultado.Cells(nFilasCont + 7, 8).Value = wsResultado.Cells(nFilasCont + 7, 8).Value + 1
                                    wsResultado.Cells(nFilasCont + 9, 8).Value = wsResultado.Cells(nFilasCont + 9, 8).Value + 1
                                Else
                                    If Cells(i, nColumnas).Value < 60 Then
                                        wsResultado.Cells(nFilasCont + 7, 10).Value = wsResultado.Cells(nFilasCont + 7, 10).Value + 1
                                        wsResultado.Cells(nFilasCont + 9, 10).Value = wsResultado.Cells(nFilasCont + 9, 10).Value + 1
                                    Else
                                        If Cells(i, nColumnas).Value < 66 Then
                                            wsResultado.Cells(nFilasCont + 7, 12).Value = wsResultado.Cells(nFilasCont + 7, 12).Value + 1
                                            wsResultado.Cells(nFilasCont + 9, 12).Value = wsResultado.Cells(nFilasCont + 9, 12).Value + 1
                                        Else
                                            wsResultado.Cells(nFilasCont + 7, 14).Value = wsResultado.Cells(nFilasCont + 7, 14).Value + 1
                                            wsResultado.Cells(nFilasCont + 9, 14).Value = wsResultado.Cells(nFilasCont + 9, 14).Value + 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                    wsResultado.Cells(nFilasCont + 7, 16).Value = wsResultado.Cells(nFilasCont + 7, 16).Value + 1
                Else
                    If Cells(i, 18).Value >= 25 Then
                        If Cells(i, nColumnas).Value < 40 Then
                            wsResultado.Cells(nFilasCont + 6, 2).Value = wsResultado.Cells(nFilasCont + 6, 2).Value + 1
                            wsResultado.Cells(nFilasCont + 9, 2).Value = wsResultado.Cells(nFilasCont + 9, 2).Value + 1
                        Else
                            If Cells(i, nColumnas).Value < 45 Then
                                wsResultado.Cells(nFilasCont + 6, 4).Value = wsResultado.Cells(nFilasCont + 6, 4).Value + 1
                                wsResultado.Cells(nFilasCont + 9, 4).Value = wsResultado.Cells(nFilasCont + 9, 4).Value + 1
                            Else
                                If Cells(i, nColumnas).Value < 50 Then
                                    wsResultado.Cells(nFilasCont + 6, 6).Value = wsResultado.Cells(nFilasCont + 6, 6).Value + 1
                                    wsResultado.Cells(nFilasCont + 9, 6).Value = wsResultado.Cells(nFilasCont + 9, 6).Value + 1
                                Else
                                    If Cells(i, nColumnas).Value < 55 Then
                                        wsResultado.Cells(nFilasCont + 6, 8).Value = wsResultado.Cells(nFilasCont + 6, 8).Value + 1
                                        wsResultado.Cells(nFilasCont + 9, 8).Value = wsResultado.Cells(nFilasCont + 9, 8).Value + 1
                                    Else
                                        If Cells(i, nColumnas).Value < 60 Then
                                            wsResultado.Cells(nFilasCont + 6, 10).Value = wsResultado.Cells(nFilasCont + 6, 10).Value + 1
                                            wsResultado.Cells(nFilasCont + 9, 10).Value = wsResultado.Cells(nFilasCont + 9, 10).Value + 1
                                        Else
                                            If Cells(i, nColumnas).Value < 66 Then
                                                wsResultado.Cells(nFilasCont + 6, 12).Value = wsResultado.Cells(nFilasCont + 6, 12).Value + 1
                                                wsResultado.Cells(nFilasCont + 9, 12).Value = wsResultado.Cells(nFilasCont + 9, 12).Value + 1
                                            Else
                                                wsResultado.Cells(nFilasCont + 6, 14).Value = wsResultado.Cells(nFilasCont + 6, 14).Value + 1
                                                wsResultado.Cells(nFilasCont + 9, 14).Value = wsResultado.Cells(nFilasCont + 9, 14).Value + 1
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        wsResultado.Cells(nFilasCont + 6, 16).Value = wsResultado.Cells(nFilasCont + 6, 16).Value + 1
                    Else
                        If Cells(i, 18).Value >= 20 Then
                            If Cells(i, nColumnas).Value < 40 Then
                                wsResultado.Cells(nFilasCont + 5, 2).Value = wsResultado.Cells(nFilasCont + 5, 2).Value + 1
                                wsResultado.Cells(nFilasCont + 9, 2).Value = wsResultado.Cells(nFilasCont + 9, 2).Value + 1
                            Else
                                If Cells(i, nColumnas).Value < 45 Then
                                    wsResultado.Cells(nFilasCont + 5, 4).Value = wsResultado.Cells(nFilasCont + 5, 4).Value + 1
                                    wsResultado.Cells(nFilasCont + 9, 4).Value = wsResultado.Cells(nFilasCont + 9, 4).Value + 1
                                Else
                                    If Cells(i, nColumnas).Value < 50 Then
                                        wsResultado.Cells(nFilasCont + 5, 6).Value = wsResultado.Cells(nFilasCont + 5, 6).Value + 1
                                        wsResultado.Cells(nFilasCont + 9, 6).Value = wsResultado.Cells(nFilasCont + 9, 6).Value + 1
                                    Else
                                        If Cells(i, nColumnas).Value < 55 Then
                                            wsResultado.Cells(nFilasCont + 5, 8).Value = wsResultado.Cells(nFilasCont + 5, 8).Value + 1
                                            wsResultado.Cells(nFilasCont + 9, 8).Value = wsResultado.Cells(nFilasCont + 9, 8).Value + 1
                                        Else
                                            If Cells(i, nColumnas).Value < 60 Then
                                                wsResultado.Cells(nFilasCont + 5, 10).Value = wsResultado.Cells(nFilasCont + 5, 10).Value + 1
                                                wsResultado.Cells(nFilasCont + 9, 10).Value = wsResultado.Cells(nFilasCont + 9, 10).Value + 1
                                            Else
                                                If Cells(i, nColumnas).Value < 66 Then
                                                    wsResultado.Cells(nFilasCont + 5, 12).Value = wsResultado.Cells(nFilasCont + 5, 12).Value + 1
                                                    wsResultado.Cells(nFilasCont + 9, 12).Value = wsResultado.Cells(nFilasCont + 9, 12).Value + 1
                                                Else
                                                    wsResultado.Cells(nFilasCont + 5, 14).Value = wsResultado.Cells(nFilasCont + 5, 14).Value + 1
                                                    wsResultado.Cells(nFilasCont + 9, 14).Value = wsResultado.Cells(nFilasCont + 9, 14).Value + 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            wsResultado.Cells(nFilasCont + 5, 16).Value = wsResultado.Cells(nFilasCont + 5, 16).Value + 1
                        Else
                            If Cells(i, 18).Value >= 15 Then
                                If Cells(i, nColumnas).Value < 40 Then
                                    wsResultado.Cells(nFilasCont + 4, 2).Value = wsResultado.Cells(nFilasCont + 4, 2).Value + 1
                                    wsResultado.Cells(nFilasCont + 9, 2).Value = wsResultado.Cells(nFilasCont + 9, 2).Value + 1
                                Else
                                    If Cells(i, nColumnas).Value < 45 Then
                                        wsResultado.Cells(nFilasCont + 4, 4).Value = wsResultado.Cells(nFilasCont + 4, 4).Value + 1
                                        wsResultado.Cells(nFilasCont + 9, 4).Value = wsResultado.Cells(nFilasCont + 9, 4).Value + 1
                                    Else
                                        If Cells(i, nColumnas).Value < 50 Then
                                            wsResultado.Cells(nFilasCont + 4, 6).Value = wsResultado.Cells(nFilasCont + 4, 6).Value + 1
                                            wsResultado.Cells(nFilasCont + 9, 6).Value = wsResultado.Cells(nFilasCont + 9, 6).Value + 1
                                        Else
                                            If Cells(i, nColumnas).Value < 55 Then
                                                wsResultado.Cells(nFilasCont + 4, 8).Value = wsResultado.Cells(nFilasCont + 4, 8).Value + 1
                                                wsResultado.Cells(nFilasCont + 9, 8).Value = wsResultado.Cells(nFilasCont + 9, 8).Value + 1
                                            Else
                                                If Cells(i, nColumnas).Value < 60 Then
                                                    wsResultado.Cells(nFilasCont + 4, 10).Value = wsResultado.Cells(nFilasCont + 4, 10).Value + 1
                                                    wsResultado.Cells(nFilasCont + 9, 10).Value = wsResultado.Cells(nFilasCont + 9, 10).Value + 1
                                                Else
                                                    If Cells(i, nColumnas).Value < 66 Then
                                                        wsResultado.Cells(nFilasCont + 4, 12).Value = wsResultado.Cells(nFilasCont + 4, 12).Value + 1
                                                        wsResultado.Cells(nFilasCont + 9, 12).Value = wsResultado.Cells(nFilasCont + 9, 12).Value + 1
                                                    Else
                                                        wsResultado.Cells(nFilasCont + 4, 14).Value = wsResultado.Cells(nFilasCont + 4, 14).Value + 1
                                                        wsResultado.Cells(nFilasCont + 9, 14).Value = wsResultado.Cells(nFilasCont + 9, 14).Value + 1
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                wsResultado.Cells(nFilasCont + 4, 16).Value = wsResultado.Cells(nFilasCont + 4, 16).Value + 1
                            Else
                                If Cells(i, 18).Value >= 10 Then
                                    If Cells(i, nColumnas).Value < 40 Then
                                        wsResultado.Cells(nFilasCont + 3, 2).Value = wsResultado.Cells(nFilasCont + 3, 2).Value + 1
                                        wsResultado.Cells(nFilasCont + 9, 2).Value = wsResultado.Cells(nFilasCont + 9, 2).Value + 1
                                    Else
                                        If Cells(i, nColumnas).Value < 45 Then
                                            wsResultado.Cells(nFilasCont + 3, 4).Value = wsResultado.Cells(nFilasCont + 3, 4).Value + 1
                                            wsResultado.Cells(nFilasCont + 9, 4).Value = wsResultado.Cells(nFilasCont + 9, 4).Value + 1
                                        Else
                                            If Cells(i, nColumnas).Value < 50 Then
                                                wsResultado.Cells(nFilasCont + 3, 6).Value = wsResultado.Cells(nFilasCont + 3, 6).Value + 1
                                                wsResultado.Cells(nFilasCont + 9, 6).Value = wsResultado.Cells(nFilasCont + 9, 6).Value + 1
                                            Else
                                                If Cells(i, nColumnas).Value < 55 Then
                                                    wsResultado.Cells(nFilasCont + 3, 8).Value = wsResultado.Cells(nFilasCont + 3, 8).Value + 1
                                                    wsResultado.Cells(nFilasCont + 9, 8).Value = wsResultado.Cells(nFilasCont + 9, 8).Value + 1
                                                Else
                                                    If Cells(i, nColumnas).Value < 60 Then
                                                        wsResultado.Cells(nFilasCont + 3, 10).Value = wsResultado.Cells(nFilasCont + 3, 10).Value + 1
                                                        wsResultado.Cells(nFilasCont + 9, 10).Value = wsResultado.Cells(nFilasCont + 9, 10).Value + 1
                                                    Else
                                                        If Cells(i, nColumnas).Value < 66 Then
                                                            wsResultado.Cells(nFilasCont + 3, 12).Value = wsResultado.Cells(nFilasCont + 3, 12).Value + 1
                                                            wsResultado.Cells(nFilasCont + 9, 12).Value = wsResultado.Cells(nFilasCont + 9, 12).Value + 1
                                                        Else
                                                            wsResultado.Cells(nFilasCont + 3, 14).Value = wsResultado.Cells(nFilasCont + 3, 14).Value + 1
                                                            wsResultado.Cells(nFilasCont + 9, 14).Value = wsResultado.Cells(nFilasCont + 9, 14).Value + 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                    wsResultado.Cells(nFilasCont + 3, 16).Value = wsResultado.Cells(nFilasCont + 3, 16).Value + 1
                                Else
                                    If Cells(i, nColumnas).Value < 40 Then
                                        wsResultado.Cells(nFilasCont + 2, 2).Value = wsResultado.Cells(nFilasCont + 2, 2).Value + 1
                                        wsResultado.Cells(nFilasCont + 9, 2).Value = wsResultado.Cells(nFilasCont + 9, 2).Value + 1
                                    Else
                                        If Cells(i, nColumnas).Value < 45 Then
                                            wsResultado.Cells(nFilasCont + 2, 4).Value = wsResultado.Cells(nFilasCont + 2, 4).Value + 1
                                            wsResultado.Cells(nFilasCont + 9, 4).Value = wsResultado.Cells(nFilasCont + 9, 4).Value + 1
                                        Else
                                            If Cells(i, nColumnas).Value < 50 Then
                                                wsResultado.Cells(nFilasCont + 2, 6).Value = wsResultado.Cells(nFilasCont + 2, 6).Value + 1
                                                wsResultado.Cells(nFilasCont + 9, 6).Value = wsResultado.Cells(nFilasCont + 9, 6).Value + 1
                                            Else
                                                If Cells(i, nColumnas).Value < 55 Then
                                                    wsResultado.Cells(nFilasCont + 2, 8).Value = wsResultado.Cells(nFilasCont + 2, 8).Value + 1
                                                    wsResultado.Cells(nFilasCont + 9, 8).Value = wsResultado.Cells(nFilasCont + 9, 8).Value + 1
                                                Else
                                                    If Cells(i, nColumnas).Value < 60 Then
                                                        wsResultado.Cells(nFilasCont + 2, 10).Value = wsResultado.Cells(nFilasCont + 2, 10).Value + 1
                                                        wsResultado.Cells(nFilasCont + 9, 10).Value = wsResultado.Cells(nFilasCont + 9, 10).Value + 1
                                                    Else
                                                        If Cells(i, nColumnas).Value < 66 Then
                                                            wsResultado.Cells(nFilasCont + 2, 12).Value = wsResultado.Cells(nFilasCont + 2, 12).Value + 1
                                                            wsResultado.Cells(nFilasCont + 9, 12).Value = wsResultado.Cells(nFilasCont + 9, 12).Value + 1
                                                        Else
                                                            wsResultado.Cells(nFilasCont + 2, 14).Value = wsResultado.Cells(nFilasCont + 2, 14).Value + 1
                                                            wsResultado.Cells(nFilasCont + 9, 14).Value = wsResultado.Cells(nFilasCont + 9, 14).Value + 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                    wsResultado.Cells(nFilasCont + 2, 16).Value = wsResultado.Cells(nFilasCont + 2, 16).Value + 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                'es cpto 1
                cantidad = cantidad + 1
                If Cells(i - 1, 12).Value <> Cells(i, 12).Value And Cells(i + 1, 12).Value <> Cells(i, 12).Value Then
                    'el dni no tiene cpto 200
                    If Cells(i, nColumnas).Value < 40 Then
                        wsResultado.Cells(nFilasCont + 8, 2).Value = wsResultado.Cells(nFilasCont + 8, 2).Value + 1
                        wsResultado.Cells(nFilasCont + 9, 2).Value = wsResultado.Cells(nFilasCont + 9, 2).Value + 1
                    Else
                        If Cells(i, nColumnas).Value < 45 Then
                            wsResultado.Cells(nFilasCont + 8, 4).Value = wsResultado.Cells(nFilasCont + 8, 4).Value + 1
                            wsResultado.Cells(nFilasCont + 9, 4).Value = wsResultado.Cells(nFilasCont + 9, 4).Value + 1
                        Else
                            If Cells(i, nColumnas).Value < 50 Then
                                wsResultado.Cells(nFilasCont + 8, 6).Value = wsResultado.Cells(nFilasCont + 8, 6).Value + 1
                                wsResultado.Cells(nFilasCont + 9, 6).Value = wsResultado.Cells(nFilasCont + 9, 6).Value + 1
                            Else
                                If Cells(i, nColumnas).Value < 55 Then
                                    wsResultado.Cells(nFilasCont + 8, 8).Value = wsResultado.Cells(nFilasCont + 8, 8).Value + 1
                                    wsResultado.Cells(nFilasCont + 9, 8).Value = wsResultado.Cells(nFilasCont + 9, 8).Value + 1
                                Else
                                    If Cells(i, nColumnas).Value < 60 Then
                                        wsResultado.Cells(nFilasCont + 8, 10).Value = wsResultado.Cells(nFilasCont + 8, 10).Value + 1
                                        wsResultado.Cells(nFilasCont + 9, 10).Value = wsResultado.Cells(nFilasCont + 9, 10).Value + 1
                                    Else
                                        If Cells(i, nColumnas).Value < 66 Then
                                            wsResultado.Cells(nFilasCont + 8, 12).Value = wsResultado.Cells(nFilasCont + 8, 12).Value + 1
                                            wsResultado.Cells(nFilasCont + 9, 12).Value = wsResultado.Cells(nFilasCont + 9, 12).Value + 1
                                        Else
                                            wsResultado.Cells(nFilasCont + 8, 14).Value = wsResultado.Cells(nFilasCont + 8, 14).Value + 1
                                            wsResultado.Cells(nFilasCont + 9, 14).Value = wsResultado.Cells(nFilasCont + 9, 14).Value + 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                    wsResultado.Cells(nFilasCont + 8, 16).Value = wsResultado.Cells(nFilasCont + 8, 16).Value + 1
                End If
            End If
        Else
            'Calcular %
            wsResultado.Cells(nFilasCont + 9, 16).Value = cantidad
            For m = 2 To 16 Step 2
                For j = 2 To 9
                    wsResultado.Cells(nFilasCont + j, m + 1).Value = wsResultado.Cells(nFilasCont + j, m).Value / cantidad
                Next j
            Next m
            
            nFilasCont = nFilasCont + 11
            jur = Cells(i, 8).Value
            cantidad = 0
            cobran = 0
            
            wsResultado.Cells(nFilasCont, 1).Value = "JUR: " & jur
            temp = "A" & nFilasCont
            wsResultado.Range(temp).Font.Bold = True
    
            For j = 2 To 9
                For m = 2 To 12 Step 2
                wsResultado.Cells(nFilasCont + j, m).Value = 0
                Next m
            Next j
            wsResultado.Cells(nFilasCont + 1, 1).Value = "ANTIG\EDAD"
            wsResultado.Cells(nFilasCont + 2, 1).Value = ".1 - 9"
            wsResultado.Cells(nFilasCont + 3, 1).Value = ".10 - 14"
            wsResultado.Cells(nFilasCont + 4, 1).Value = ".15 - 19"
            wsResultado.Cells(nFilasCont + 5, 1).Value = ".20 - 24"
            wsResultado.Cells(nFilasCont + 6, 1).Value = ".25 - 29"
            wsResultado.Cells(nFilasCont + 7, 1).Value = ".30 o más"
            wsResultado.Cells(nFilasCont + 8, 1).Value = "No cobran"
            wsResultado.Cells(nFilasCont + 9, 1).Value = "TOTAL"
            wsResultado.Cells(nFilasCont + 1, 2).Value = "Menor 40"
            wsResultado.Cells(nFilasCont + 1, 4).Value = "40 - 44"
            wsResultado.Cells(nFilasCont + 1, 6).Value = "45 - 49"
            wsResultado.Cells(nFilasCont + 1, 8).Value = "50 - 54"
            wsResultado.Cells(nFilasCont + 1, 10).Value = "55 - 59"
            wsResultado.Cells(nFilasCont + 1, 12).Value = "60 - 65"
            wsResultado.Cells(nFilasCont + 1, 14).Value = "Mayor 65"
            wsResultado.Cells(nFilasCont + 1, 16).Value = "TOTAL"
            
            Application.DisplayAlerts = False
            fila = nFilasCont + 1
            wsResultado.Range("B" & fila & ":C" & fila).Merge
            wsResultado.Range("B" & fila & ":C" & fila).HorizontalAlignment = xlCenter
            wsResultado.Range("B" & fila & ":C" & fila).Font.Bold = True
            wsResultado.Range("D" & fila & ":E" & fila).Merge
            wsResultado.Range("D" & fila & ":E" & fila).HorizontalAlignment = xlCenter
            wsResultado.Range("D" & fila & ":E" & fila).Font.Bold = True
            wsResultado.Range("F" & fila & ":G" & fila).Merge
            wsResultado.Range("F" & fila & ":G" & fila).HorizontalAlignment = xlCenter
            wsResultado.Range("F" & fila & ":G" & fila).Font.Bold = True
            wsResultado.Range("H" & fila & ":I" & fila).Merge
            wsResultado.Range("H" & fila & ":I" & fila).HorizontalAlignment = xlCenter
            wsResultado.Range("H" & fila & ":I" & fila).Font.Bold = True
            wsResultado.Range("J" & fila & ":K" & fila).Merge
            wsResultado.Range("J" & fila & ":K" & fila).HorizontalAlignment = xlCenter
            wsResultado.Range("J" & fila & ":K" & fila).Font.Bold = True
            wsResultado.Range("L" & fila & ":M" & fila).Merge
            wsResultado.Range("L" & fila & ":M" & fila).HorizontalAlignment = xlCenter
            wsResultado.Range("L" & fila & ":M" & fila).Font.Bold = True
            wsResultado.Range("N" & fila & ":O" & fila).Merge
            wsResultado.Range("N" & fila & ":O" & fila).HorizontalAlignment = xlCenter
            wsResultado.Range("N" & fila & ":O" & fila).Font.Bold = True
            wsResultado.Range("P" & fila & ":Q" & fila).Merge
            wsResultado.Range("P" & fila & ":Q" & fila).HorizontalAlignment = xlCenter
            wsResultado.Range("P" & fila & ":Q" & fila).Font.Bold = True
            wsResultado.Range("R" & fila & ":S" & fila).Merge
            wsResultado.Range("R" & fila & ":S" & fila).HorizontalAlignment = xlCenter
            wsResultado.Range("R" & fila & ":S" & fila).Font.Bold = True
            Application.DisplayAlerts = True
            
            i = i - 1
        End If
        
    Next i
    
    'Calcular %
    wsResultado.Cells(nFilasCont + 9, 16).Value = cantidad
    For m = 2 To 16 Step 2
        For j = 2 To 9
            wsResultado.Cells(nFilasCont + j, m + 1).Value = wsResultado.Cells(nFilasCont + j, m).Value / cantidad
        Next j
    Next m
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub


Sub Informe_Direct_JefDepto()
    Dim wbContenido As Workbook, _
        wsResultado As Excel.Worksheet, _
        wsDetalle As Excel.Worksheet, _
        wsContenido As Excel.Worksheet
    Dim nFilas As Double
    Dim nColumnas As Integer
    Dim rango As Range
    Dim nFilasCont As Double
    Dim rangoCont As Range
    Dim i As Long
    Dim j As Long
    Dim m As Integer
    Dim nFilasDetalle As Long

    
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
    
    Worksheets.Add
    ActiveSheet.Name = "Detalle"
    Application.DisplayAlerts = True
    
    Set wsResultado = Worksheets("Resultado")
    Set wsDetalle = Worksheets("Detalle")
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
    
    
    nFilasDetalle = 1
    wsDetalle.Cells(nFilasDetalle, 1).Value = "JUR"
    wsDetalle.Cells(nFilasDetalle, 2).Value = "ESC"
    wsDetalle.Cells(nFilasDetalle, 3).Value = "DOC"
    wsDetalle.Cells(nFilasDetalle, 4).Value = "NOMBRE"
    wsDetalle.Cells(nFilasDetalle, 5).Value = "CEIC"
    wsDetalle.Cells(nFilasDetalle, 6).Value = "OFICINA"
    wsDetalle.Cells(nFilasDetalle, 7).Value = "ANEXO"
    wsDetalle.Cells(nFilasDetalle, 8).Value = "EDAD"
    wsDetalle.Cells(nFilasDetalle, 9).Value = "ANTIGÜEDAD"
    wsDetalle.Cells(nFilasDetalle, 10).Value = "CARGO"
    wsDetalle.Cells(nFilasDetalle, 11).Value = "SITUACIÓN"
    wsDetalle.Cells(nFilasDetalle, 12).Value = "JUR"
    wsDetalle.Cells(nFilasDetalle, 13).Value = "PREF"
    nFilasDetalle = 2
    
    
    For i = 2 To nFilas
        If Cells(i, 4).Value = 1 Then
            valorDoc = Cells(i, 12).Value
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
                
                bandera = False
                
                If Cells(i + 1, 4).Value = 200 Then
                    i = i + 1
                Else
                    bandera = True
                End If
                
                wsDetalle.Cells(nFilasDetalle, 1).Value = Cells(i, 8).Value
                wsDetalle.Cells(nFilasDetalle, 2).Value = Cells(i, 9).Value
                wsDetalle.Cells(nFilasDetalle, 3).Value = Cells(i, 12).Value
                wsDetalle.Cells(nFilasDetalle, 4).Value = Cells(i, 14).Value
                wsDetalle.Cells(nFilasDetalle, 5).Value = Cells(i, 15).Value
                wsDetalle.Cells(nFilasDetalle, 6).Value = Cells(i, 20).Value
                wsDetalle.Cells(nFilasDetalle, 7).Value = Cells(i, 21).Value
                wsDetalle.Cells(nFilasDetalle, 8).Value = Cells(i, 29).Value
                wsDetalle.Cells(nFilasDetalle, 9).Value = Cells(i, 18).Value
                wsDetalle.Cells(nFilasDetalle, 10).Value = wsContenido.Cells(j, 9).Value
                wsDetalle.Cells(nFilasDetalle, 11).Value = wsContenido.Cells(j, 10).Value
                wsDetalle.Cells(nFilasDetalle, 12).Value = wsContenido.Cells(j, 1).Value
                wsDetalle.Cells(nFilasDetalle, 13).Value = wsContenido.Cells(j, 2).Value
                nFilasDetalle = nFilasDetalle + 1
                
                tempFila = 0
                tempColum = 0
                If wsContenido.Cells(j, 7).Value = 1015 Then
                    tempFila = 47
                Else
                    If wsContenido.Cells(j, 7).Value = 1040 Then
                        tempFila = 33
                    Else
                        tempFila = 19
                    End If
                End If
                
                If bandera Then
                    tempFila = tempFila + 7
                Else
                    If Cells(i, 18).Value < 10 Then
                        tempFila = tempFila + 1
                    Else
                        If Cells(i, 18).Value < 15 Then
                            tempFila = tempFila + 2
                        Else
                            If Cells(i, 18).Value < 20 Then
                                tempFila = tempFila + 3
                            Else
                                If Cells(i, 18).Value < 25 Then
                                    tempFila = tempFila + 4
                                Else
                                    If Cells(i, 18).Value < 30 Then
                                        tempFila = tempFila + 5
                                    Else
                                        If Cells(i, 18).Value > 29 Then
                                            tempFila = tempFila + 6
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                
                If Cells(i, 29).Value < 40 Then
                    tempColum = 4
                Else
                    If Cells(i, 29).Value < 45 Then
                        tempColum = 6
                    Else
                        If Cells(i, 29).Value < 50 Then
                            tempColum = 8
                        Else
                            If Cells(i, 29).Value < 55 Then
                                tempColum = 10
                            Else
                                If Cells(i, 29).Value < 60 Then
                                    tempColum = 12
                                Else
                                    If Cells(i, 29).Value < 66 Then
                                        tempColum = 14
                                    Else
                                        If Cells(i, 29).Value > 65 Then
                                            tempColum = 16
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                
                If wsContenido.Cells(j, 10).Value = "Subroga Cargo" Then
                    tempColum = tempColum + 1
                End If
                
                If wsResultado.Cells(tempFila, tempColum).Value = "" Then
                    wsResultado.Cells(tempFila, tempColum).Value = 1
                Else
                    wsResultado.Cells(tempFila, tempColum).Value = wsResultado.Cells(tempFila, tempColum).Value + 1
                End If
                
            End If
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub
