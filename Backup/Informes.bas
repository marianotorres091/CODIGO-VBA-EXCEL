Attribute VB_Name = "Module1"
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


Sub Informe_Complemento()
    Dim wsResultado As Excel.Worksheet
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim nFilasCont As Long
    Dim valorDoc As String
    Dim jur As Integer
    Dim cantidad As Long
    Dim cobran As Long
    Dim fechaDesde As Date
    Dim fechaHasta As Date
    Dim fechaTemp As Date
    
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
    
    wsResultado.Cells(1, 1).Value = "Fecha Desde:"
    fechaDesde = Cells(2, 16).Value
    wsResultado.Cells(1, 3).Value = "Fecha Hasta:"
    fechaHasta = Cells(2, 16).Value
    
    jur = Cells(2, 1).Value
    
    nFilasCont = 2
    wsResultado.Cells(nFilasCont, 1).Value = "JUR: " & jur
    'wsResultado.Range("A1").Font.Bold = True
    nFilasCont = nFilasCont + 1
    wsResultado.Cells(nFilasCont, 1).Value = "Complemento"
    wsResultado.Cells(nFilasCont, 2).Value = "Registrados"
    wsResultado.Cells(nFilasCont, 3).Value = "Importe"
    wsResultado.Cells(nFilasCont, 4).Value = "Autorizados"
    wsResultado.Cells(nFilasCont, 5).Value = "Importe"
    wsResultado.Cells(nFilasCont, 6).Value = "Total"
    wsResultado.Cells(nFilasCont, 7).Value = "Importe"
    wsResultado.Cells(nFilasCont + 1, 1).Value = "Bon. Ag. Est. Sanitarios"
    wsResultado.Cells(nFilasCont + 2, 1).Value = "Bon. Esp. Ley 6602"
    wsResultado.Cells(nFilasCont + 3, 1).Value = "Dedicación Exclusiva"
    wsResultado.Cells(nFilasCont + 4, 1).Value = "Ded. Excl. Dcto. 2441/15"
    wsResultado.Cells(nFilasCont + 5, 1).Value = "Riesgo de Vida"
    wsResultado.Cells(nFilasCont + 6, 1).Value = "Tít. Secundario"
    wsResultado.Cells(nFilasCont + 7, 1).Value = "Tít. Universitario"
    wsResultado.Cells(nFilasCont + 8, 1).Value = "Zona Desfavorable"
    wsResultado.Cells(nFilasCont + 9, 1).Value = "Riesgo de Salud"
    wsResultado.Cells(nFilasCont + 10, 1).Value = "Total"
    For j = 1 To 10
        For m = 2 To 7
            wsResultado.Cells(nFilasCont + j, m).Value = 0
        Next m
    Next j
    
    For i = 2 To nFilas
        fechaTemp = Cells(i, 16).Value
        If fechaTemp < fechaDesde Then
            fechaDesde = fechaTemp
        Else
            If fechaTemp > fechaHasta Then
                fechaHasta = fechaTemp
            End If
        End If
        If jur = Cells(i, 1).Value Then
            If Cells(i, 8).Value = 4 Then
                temp = 7
            Else
                If Cells(i, 8).Value = 3 Then
                    temp = 6
                Else
                    If Cells(i, 8).Value = 5 Then
                        temp = 9
                    Else
                        If Cells(i, 8).Value = 32 Then
                            temp = 1
                        Else
                            If Cells(i, 8).Value = 2 Then
                                temp = 3
                            Else
                                If Cells(i, 8).Value = 46 Then
                                    temp = 4
                                Else
                                    If Cells(i, 8).Value = 23 Then
                                        temp = 2
                                    Else
                                        If Cells(i, 8).Value = 6 Then
                                            temp = 5
                                        Else
                                            temp = 8
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            
            If Cells(i, 21).Value = "Registrado" Then
                wsResultado.Cells(nFilasCont + temp, 2).Value = wsResultado.Cells(nFilasCont + temp, 2).Value + 1
                wsResultado.Cells(nFilasCont + temp, 3).Value = wsResultado.Cells(nFilasCont + temp, 3).Value + Cells(i, 24).Value
            Else
                wsResultado.Cells(nFilasCont + temp, 4).Value = wsResultado.Cells(nFilasCont + temp, 4).Value + 1
                wsResultado.Cells(nFilasCont + temp, 5).Value = wsResultado.Cells(nFilasCont + temp, 5).Value + Cells(i, 24).Value
            End If
        Else
            'sumar totales
            For j = 1 To 9
                wsResultado.Cells(nFilasCont + j, 6).Value = wsResultado.Cells(nFilasCont + j, 2).Value + wsResultado.Cells(nFilasCont + j, 4).Value
                wsResultado.Cells(nFilasCont + j, 7).Value = wsResultado.Cells(nFilasCont + j, 3).Value + wsResultado.Cells(nFilasCont + j, 5).Value
                
                wsResultado.Cells(nFilasCont + 10, 2).Value = wsResultado.Cells(nFilasCont + j, 2).Value + wsResultado.Cells(nFilasCont + 10, 2).Value
                wsResultado.Cells(nFilasCont + 10, 3).Value = wsResultado.Cells(nFilasCont + j, 3).Value + wsResultado.Cells(nFilasCont + 10, 3).Value
                wsResultado.Cells(nFilasCont + 10, 4).Value = wsResultado.Cells(nFilasCont + j, 4).Value + wsResultado.Cells(nFilasCont + 10, 4).Value
                wsResultado.Cells(nFilasCont + 10, 5).Value = wsResultado.Cells(nFilasCont + j, 5).Value + wsResultado.Cells(nFilasCont + 10, 5).Value
                
                wsResultado.Cells(nFilasCont + 10, 6).Value = wsResultado.Cells(nFilasCont + 10, 6).Value + wsResultado.Cells(nFilasCont + j, 6).Value
                wsResultado.Cells(nFilasCont + 10, 7).Value = wsResultado.Cells(nFilasCont + 10, 7).Value + wsResultado.Cells(nFilasCont + j, 7).Value
            Next j
            
            jur = Cells(i, 1).Value
            nFilasCont = nFilasCont + 11
            wsResultado.Cells(nFilasCont, 1).Value = "JUR: " & jur
            nFilasCont = nFilasCont + 1
            wsResultado.Cells(nFilasCont, 1).Value = "Complemento"
            wsResultado.Cells(nFilasCont, 2).Value = "Registrados"
            wsResultado.Cells(nFilasCont, 3).Value = "Importe"
            wsResultado.Cells(nFilasCont, 4).Value = "Autorizados"
            wsResultado.Cells(nFilasCont, 5).Value = "Importe"
            wsResultado.Cells(nFilasCont, 6).Value = "Total"
            wsResultado.Cells(nFilasCont, 7).Value = "Importe"
            wsResultado.Cells(nFilasCont + 1, 1).Value = "Bon. Ag. Est. Sanitarios"
            wsResultado.Cells(nFilasCont + 2, 1).Value = "Bon. Esp. Ley 6602"
            wsResultado.Cells(nFilasCont + 3, 1).Value = "Dedicación Exclusiva"
            wsResultado.Cells(nFilasCont + 4, 1).Value = "Ded. Excl. Dcto. 2441/15"
            wsResultado.Cells(nFilasCont + 5, 1).Value = "Riesgo de Vida"
            wsResultado.Cells(nFilasCont + 6, 1).Value = "Tít. Secundario"
            wsResultado.Cells(nFilasCont + 7, 1).Value = "Tít. Universitario"
            wsResultado.Cells(nFilasCont + 8, 1).Value = "Zona Desfavorable"
            wsResultado.Cells(nFilasCont + 9, 1).Value = "Riesgo de Salud"
            wsResultado.Cells(nFilasCont + 10, 1).Value = "Total"
            For j = 1 To 10
                For m = 2 To 7
                    wsResultado.Cells(nFilasCont + j, m).Value = 0
                Next m
            Next j
            i = i - 1
        End If
    Next i
    
    For j = 1 To 9
        wsResultado.Cells(nFilasCont + j, 6).Value = wsResultado.Cells(nFilasCont + j, 2).Value + wsResultado.Cells(nFilasCont + j, 4).Value
        wsResultado.Cells(nFilasCont + j, 7).Value = wsResultado.Cells(nFilasCont + j, 3).Value + wsResultado.Cells(nFilasCont + j, 5).Value
        
        wsResultado.Cells(nFilasCont + 10, 2).Value = wsResultado.Cells(nFilasCont + j, 2).Value + wsResultado.Cells(nFilasCont + 10, 2).Value
        wsResultado.Cells(nFilasCont + 10, 3).Value = wsResultado.Cells(nFilasCont + j, 3).Value + wsResultado.Cells(nFilasCont + 10, 3).Value
        wsResultado.Cells(nFilasCont + 10, 4).Value = wsResultado.Cells(nFilasCont + j, 4).Value + wsResultado.Cells(nFilasCont + 10, 4).Value
        wsResultado.Cells(nFilasCont + 10, 5).Value = wsResultado.Cells(nFilasCont + j, 5).Value + wsResultado.Cells(nFilasCont + 10, 5).Value
        
        wsResultado.Cells(nFilasCont + 10, 6).Value = wsResultado.Cells(nFilasCont + 10, 6).Value + wsResultado.Cells(nFilasCont + j, 6).Value
        wsResultado.Cells(nFilasCont + 10, 7).Value = wsResultado.Cells(nFilasCont + 10, 7).Value + wsResultado.Cells(nFilasCont + j, 7).Value
    Next j
    
    wsResultado.Cells(1, 2).Value = fechaDesde
    wsResultado.Cells(1, 4).Value = fechaHasta
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub



Sub Totales_Cpto_Jur()
    Dim nFilas As Long
    Dim filaTotal As Long
    Dim rango As Range
    Dim wsTotal As Excel.Worksheet
    Dim i As Long
    Dim ultJur As Integer
    Dim ultCpto As Integer
    Dim total As Double
    
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Total Cpto"
    Application.DisplayAlerts = True
    
    Set wsTotal = Worksheets("Total Cpto")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    MsgBox "Debe estar ordenado por JUR + CPTO.", , "Atención!!"
    
    'Encabezado Hoja Totales
    wsTotal.Cells(1, 1).Value = "JUR"
    wsTotal.Cells(1, 2).Value = "CPTO"
    wsTotal.Cells(1, 3).Value = "Descripción"
    wsTotal.Cells(1, 4).Value = "Importe"
    filaTotal = 2
    
    ultJur = Cells(2, 8).Value
    ultCpto = Cells(2, 4).Value
    descripcion = Cells(2, 5).Value
    importe = 0
    total = 0
    
    For i = 2 To nFilas
        If Cells(i, 4).Value < 340 And Cells(i, 4).Value <> 267 And Cells(i, 4).Value <> 272 Then
            If ultCpto = Cells(i, 4).Value And ultJur = Cells(i, 8).Value Then
                If Cells(i, 6).Value = 2 Then
                    importe = importe - Cells(i, 7).Value
                Else
                    importe = importe + Cells(i, 7).Value
                End If
            Else
                wsTotal.Cells(filaTotal, 1).Value = ultJur
                wsTotal.Cells(filaTotal, 2).Value = ultCpto
                wsTotal.Cells(filaTotal, 3).Value = descripcion
                wsTotal.Cells(filaTotal, 4).Value = importe
                total = total + importe
                filaTotal = filaTotal + 1
                If ultJur <> Cells(i, 8).Value Then
                    wsTotal.Cells(filaTotal, 4).Value = total
                    total = 0
                    filaTotal = filaTotal + 1
                End If
                ultCpto = Cells(i, 4).Value
                ultJur = Cells(i, 8).Value
                descripcion = Cells(i, 5).Value
                importe = 0
                i = i - 1
            End If
        End If
    Next i
    wsTotal.Cells(filaTotal, 1).Value = ultJur
    wsTotal.Cells(filaTotal, 2).Value = ultCpto
    wsTotal.Cells(filaTotal, 3).Value = descripcion
    wsTotal.Cells(filaTotal, 4).Value = importe
    wsTotal.Cells(filaTotal + 1, 4).Value = total + importe
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub


Sub Totales_Cptos()
    Dim wsResultado As Excel.Worksheet
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim nFilasResultado As Long
    Dim i As Long
    Dim filaCopia As Long
    Dim columnaCopia As Long
    Dim ultCpto As Integer
    
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Resultados"
    Application.DisplayAlerts = True
    
    Set wsResultado = Worksheets("Resultados")
    
    'Regresa el control a la hoja de origen
    Sheets("TotCon_J6_2018").Select
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Fila del encabezado Resultado
    nFilasResultado = 1
    wsResultado.Cells(nFilasResultado, 1).Value = "Cpto"
    wsResultado.Cells(nFilasResultado, 2).Value = "Descripción"
    wsResultado.Cells(nFilasResultado, 3).Value = "Enero"
    wsResultado.Cells(nFilasResultado, 4).Value = "Febrero"
    wsResultado.Cells(nFilasResultado, 5).Value = "Marzo"
    wsResultado.Cells(nFilasResultado, 6).Value = "Abril"
    wsResultado.Cells(nFilasResultado, 7).Value = "Mayo"
    wsResultado.Cells(nFilasResultado, 8).Value = "Junio"
    wsResultado.Cells(nFilasResultado, 9).Value = "Julio"
    wsResultado.Cells(nFilasResultado, 10).Value = "Agosto"
    nFilasResultado = 2 - 3
    
    ultCpto = 0
    
    For i = 2 To nFilas
        If Cells(i, 5).Value = ultCpto Then
            If Cells(i, 6).Value = 0 Then
                wsResultado.Cells(nFilasResultado, 2 + Cells(i, 7).Value).Value = wsResultado.Cells(nFilasResultado, 2 + Cells(i, 7).Value).Value + Cells(i, 2).Value
                wsResultado.Cells(nFilasResultado + 1, 2 + Cells(i, 7).Value).Value = wsResultado.Cells(nFilasResultado + 1, 2 + Cells(i, 7).Value).Value + Cells(i, 4).Value
            Else
                If Cells(i, 6).Value = 1 Then
                    wsResultado.Cells(nFilasResultado + 2, 2 + Cells(i, 7).Value).Value = wsResultado.Cells(nFilasResultado + 2, 2 + Cells(i, 7).Value).Value + Cells(i, 4).Value
                Else
                    wsResultado.Cells(nFilasResultado + 2, 2 + Cells(i, 7).Value).Value = wsResultado.Cells(nFilasResultado + 2, 2 + Cells(i, 7).Value).Value - Cells(i, 4).Value
                End If
            End If
        Else
            nFilasResultado = nFilasResultado + 3
            ultCpto = Cells(i, 5).Value
            wsResultado.Cells(nFilasResultado, 1).Value = ultCpto
            wsResultado.Cells(nFilasResultado, 2).Value = "Cantidad Liq"
            wsResultado.Cells(nFilasResultado + 1, 2).Value = "Básico"
            wsResultado.Cells(nFilasResultado + 2, 2).Value = "Ajuste"
            i = i - 1
        End If
    Next i
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub


Sub Totales_Cptos_CEIC()
    Dim wsResultado As Excel.Worksheet
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim nFilasResultado As Long
    Dim i As Long
    Dim filaCopia As Long
    Dim columnaCopia As Long
    Dim ultCpto As Integer
    
    
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
    wsResultado.Cells(nFilasResultado, 1).Value = "Cargo"
    wsResultado.Cells(nFilasResultado, 2).Value = "Descripción"
    wsResultado.Cells(nFilasResultado, 3).Value = "Cpto 1"
    wsResultado.Cells(nFilasResultado, 4).Value = "Cant"
    wsResultado.Cells(nFilasResultado, 5).Value = "Cpto 120"
    wsResultado.Cells(nFilasResultado, 6).Value = "Cant"
    wsResultado.Cells(nFilasResultado, 7).Value = "Cpto 126"
    wsResultado.Cells(nFilasResultado, 8).Value = "Cant"
    wsResultado.Cells(nFilasResultado, 9).Value = "Cpto 273"
    wsResultado.Cells(nFilasResultado, 10).Value = "Cant"
    wsResultado.Cells(nFilasResultado, 11).Value = "Cpto 274"
    wsResultado.Cells(nFilasResultado, 12).Value = "Cant"
    wsResultado.Cells(nFilasResultado, 13).Value = "Guardias"
    wsResultado.Cells(nFilasResultado, 14).Value = "Cant"
    
    ultCeic = 0
    
    
    For i = 2 To nFilas
        If Cells(i, 16).Value <> ultCeic Then
            nFilasResultado = nFilasResultado + 1
            ultCeic = Cells(i, 16).Value
            wsResultado.Cells(nFilasResultado, 1).Value = Cells(i, 16).Value
        End If
        
        If Cells(i, 11).Value = 1 Then
            wsResultado.Cells(nFilasResultado, 3).Value = wsResultado.Cells(nFilasResultado, 3).Value + Cells(i, 14).Value
            wsResultado.Cells(nFilasResultado, 4).Value = wsResultado.Cells(nFilasResultado, 4).Value + 1
        Else
            If Cells(i, 11).Value = 120 Then
                wsResultado.Cells(nFilasResultado, 5).Value = wsResultado.Cells(nFilasResultado, 5).Value + Cells(i, 14).Value
                wsResultado.Cells(nFilasResultado, 6).Value = wsResultado.Cells(nFilasResultado, 6).Value + 1
            Else
                If Cells(i, 11).Value = 126 Then
                    wsResultado.Cells(nFilasResultado, 7).Value = wsResultado.Cells(nFilasResultado, 7).Value + Cells(i, 14).Value
                    wsResultado.Cells(nFilasResultado, 8).Value = wsResultado.Cells(nFilasResultado, 8).Value + 1
                Else
                    If Cells(i, 11).Value = 273 Then
                        wsResultado.Cells(nFilasResultado, 9).Value = wsResultado.Cells(nFilasResultado, 9).Value + Cells(i, 14).Value
                        wsResultado.Cells(nFilasResultado, 10).Value = wsResultado.Cells(nFilasResultado, 10).Value + 1
                    Else
                        If Cells(i, 11).Value = 274 Then
                            wsResultado.Cells(nFilasResultado, 11).Value = wsResultado.Cells(nFilasResultado, 11).Value + Cells(i, 14).Value
                            wsResultado.Cells(nFilasResultado, 12).Value = wsResultado.Cells(nFilasResultado, 12).Value + 1
                        Else
                            wsResultado.Cells(nFilasResultado, 13).Value = wsResultado.Cells(nFilasResultado, 13).Value + Cells(i, 14).Value
                            If Cells(i - 1, 8).Value <> Cells(i - 1, 8).Value Or Cells(i - 1, 11).Value < 275 Then
                                wsResultado.Cells(nFilasResultado, 14).Value = wsResultado.Cells(nFilasResultado, 14).Value + 1
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
    Next i

    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

