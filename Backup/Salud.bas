Attribute VB_Name = "Module1"
Sub Control_PrProf()
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim i As Long
    Dim filaResultado As Integer
    Dim wsResultado As Excel.Worksheet
    Dim bandera As Boolean
    
    'Agrego la nueva hoja
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Resultados"
    Application.DisplayAlerts = True
    Set wsResultado = Worksheets("Resultados")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count
     

    wsResultado.Cells(1, 1).Value = "Agentes tratado con Pta. Tipo 1:"
    wsResultado.Cells(2, 1).Value = "Agentes con Título Secundario: T.Sec"
    wsResultado.Cells(3, 1).Value = "Agentes con Título Universitario: T.Univ"
    wsResultado.Cells(4, 1).Value = "Agentes con PPS: PPS"
    wsResultado.Cells(5, 1).Value = "Agentes con Dedicación Exclusiva: D.Excl"
    wsResultado.Cells(7, 1).Value = "Total de PPS distinto a 40%:"
    wsResultado.Cells(8, 1).Value = "Total de T. Univ menores a 20%:"
    
    wsResultado.Cells(10, 2).Value = "T.Sec"
    wsResultado.Cells(10, 3).Value = "T.Univ"
    wsResultado.Cells(10, 4).Value = "PPS"
    wsResultado.Cells(10, 5).Value = "D.Excl"
    wsResultado.Cells(10, 6).Value = "Total"
    wsResultado.Cells(11, 1).Value = "T.Sec"
    wsResultado.Cells(12, 1).Value = "T.Univ"
    wsResultado.Cells(13, 1).Value = "PPS"
    wsResultado.Cells(14, 1).Value = "D.Excl"
    wsResultado.Cells(15, 1).Value = "Total"
    
    For i = 11 To 15
        For j = 2 To 6
            wsResultado.Cells(i, j).Value = 0
        Next j
    Next i
    
    wsResultado.Cells(7, 4).Value = 0
    wsResultado.Cells(8, 4).Value = 0
    
    For i = 2 To nFilas
        If Cells(i, 1).Value = 1 Then
            wsResultado.Cells(1, 4).Value = wsResultado.Cells(1, 4).Value + 1
            bandera2 = False
            'PPS
            If Cells(i, 32).Value > 0 Then
                bandera = True
                bandera2 = True
                If Cells(i, 33).Value <> 40 Then
                    wsResultado.Cells(7, 4).Value = wsResultado.Cells(7, 4).Value + 1
                End If
                'DExcl
                If Cells(i, 34).Value > 0 Then
                    wsResultado.Cells(13, 5).Value = wsResultado.Cells(13, 5).Value + 1
                    wsResultado.Cells(14, 4).Value = wsResultado.Cells(14, 4).Value + 1
                    bandera = False
                End If
                'TSec
                If Cells(i, 36).Value > 0 Then
                    wsResultado.Cells(13, 2).Value = wsResultado.Cells(13, 2).Value + 1
                    wsResultado.Cells(11, 4).Value = wsResultado.Cells(11, 4).Value + 1
                    bandera = False
                End If
                'TUniv
                If Cells(i, 38).Value > 0 Then
                    wsResultado.Cells(13, 3).Value = wsResultado.Cells(13, 3).Value + 1
                    wsResultado.Cells(12, 4).Value = wsResultado.Cells(12, 4).Value + 1
                    bandera = False
                End If
                If bandera Then
                    wsResultado.Cells(13, 4).Value = wsResultado.Cells(13, 4).Value + 1
                End If
            End If
            'TUniv
            If Cells(i, 38).Value > 0 Then
                bandera2 = True
                bandera = True
                If Cells(i, 39).Value < 21 Then
                    wsResultado.Cells(8, 4).Value = wsResultado.Cells(8, 4).Value + 1
                End If
                'DExcl
                If Cells(i, 34).Value > 0 Then
                    wsResultado.Cells(12, 5).Value = wsResultado.Cells(12, 5).Value + 1
                    wsResultado.Cells(14, 3).Value = wsResultado.Cells(14, 3).Value + 1
                    bandera = False
                End If
                'TSec
                If Cells(i, 36).Value > 0 Then
                    wsResultado.Cells(12, 2).Value = wsResultado.Cells(12, 2).Value + 1
                    wsResultado.Cells(11, 3).Value = wsResultado.Cells(11, 3).Value + 1
                    bandera = False
                End If
                If bandera And Cells(i, 32) <= 0 Then
                    wsResultado.Cells(12, 3).Value = wsResultado.Cells(12, 3).Value + 1
                End If
            End If
            'DExcl
            If Cells(i, 34).Value > 0 Then
                bandera2 = True
                If Cells(i, 36).Value > 0 Then
                    wsResultado.Cells(14, 2).Value = wsResultado.Cells(14, 2).Value + 1
                    wsResultado.Cells(11, 5).Value = wsResultado.Cells(11, 5).Value + 1
                End If
                If Cells(i, 32).Value <= 0 And Cells(i, 36).Value <= 0 And Cells(i, 38) <= 0 Then
                    wsResultado.Cells(14, 5).Value = wsResultado.Cells(14, 5).Value + 1
                End If
            End If
            'TSec
            If Cells(i, 36) > 0 Then
                bandera2 = True
                If Cells(i, 32).Value <= 0 And Cells(i, 34).Value <= 0 And Cells(i, 38) <= 0 Then
                    wsResultado.Cells(11, 2).Value = wsResultado.Cells(11, 2).Value + 1
                End If
            End If
            If bandera2 Then
                wsResultado.Cells(15, 6).Value = wsResultado.Cells(15, 6).Value + 1
            End If
        End If
    Next i
    
    'Totales
    wsResultado.Cells(15, 2).Value = wsResultado.Cells(11, 2).Value + wsResultado.Cells(12, 2).Value + wsResultado.Cells(13, 2).Value + wsResultado.Cells(14, 2).Value
    wsResultado.Cells(15, 3).Value = wsResultado.Cells(11, 3).Value + wsResultado.Cells(12, 3).Value + wsResultado.Cells(13, 3).Value + wsResultado.Cells(14, 3).Value
    wsResultado.Cells(15, 4).Value = wsResultado.Cells(11, 4).Value + wsResultado.Cells(12, 4).Value + wsResultado.Cells(13, 4).Value + wsResultado.Cells(14, 4).Value
    
    wsResultado.Cells(11, 6).Value = wsResultado.Cells(11, 2).Value + wsResultado.Cells(11, 3).Value + wsResultado.Cells(11, 4).Value + wsResultado.Cells(11, 5).Value
    wsResultado.Cells(12, 6).Value = wsResultado.Cells(12, 2).Value + wsResultado.Cells(12, 3).Value + wsResultado.Cells(12, 4).Value + wsResultado.Cells(12, 5).Value
    wsResultado.Cells(13, 6).Value = wsResultado.Cells(13, 2).Value + wsResultado.Cells(13, 3).Value + wsResultado.Cells(13, 4).Value + wsResultado.Cells(13, 5).Value
    wsResultado.Cells(14, 6).Value = wsResultado.Cells(14, 2).Value + wsResultado.Cells(14, 3).Value + wsResultado.Cells(14, 4).Value + wsResultado.Cells(14, 5).Value
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

