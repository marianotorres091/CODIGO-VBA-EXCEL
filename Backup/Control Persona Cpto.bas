Attribute VB_Name = "Módulo1"
Sub Listar_Persona_Cpto()
    Dim nFilas As Long
    Dim rango As Range
    Dim wsTotal As Excel.Worksheet
    Dim i As Long
    Dim ultDoc As String
    Dim ultCpto As Integer
    
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Resultado"
    Application.DisplayAlerts = True
    
    Set wsTotal = Worksheets("Resultado")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    MsgBox "Debe estar ordenado por DNI.", , "Atención!!"
    
    'Encabezado Hoja Totales
    wsTotal.Cells(1, 1).Value = "JUR"
    wsTotal.Cells(1, 2).Value = "DNI"
    wsTotal.Cells(1, 3).Value = "Nombre"
    wsTotal.Cells(1, 4).Value = "Cpto 1"
    wsTotal.Cells(1, 5).Value = "Cpto 100"
    wsTotal.Cells(1, 6).Value = "Cpto 246"
    wsTotal.Cells(1, 7).Value = "Ceic"
    wsTotal.Cells(1, 8).Value = "PtaTipo"
    wsTotal.Cells(1, 9).Value = "Apartado"
    wsTotal.Cells(1, 10).Value = "Categoría"
    filaTotal = 1
    
    ultDoc = 0
    
    For i = 2 To nFilas
        If ultDoc = Cells(i, 12).Value Then
            If Cells(i, 4).Value = 1 Then
                wsTotal.Cells(filaTotal, 4).Value = wsTotal.Cells(filaTotal, 4).Value + Cells(i, 7).Value
            Else
                If Cells(i, 4).Value = 100 Then
                    wsTotal.Cells(filaTotal, 5).Value = wsTotal.Cells(filaTotal, 5).Value + Cells(i, 7).Value
                Else
                    If Cells(i, 4).Value = 246 Then
                        wsTotal.Cells(filaTotal, 6).Value = wsTotal.Cells(filaTotal, 6).Value + Cells(i, 7).Value
                    End If
                End If
            End If
        Else
            filaTotal = filaTotal + 1
            wsTotal.Cells(filaTotal, 1).Value = Cells(i, 8).Value
            wsTotal.Cells(filaTotal, 2).Value = Cells(i, 12).Value
            wsTotal.Cells(filaTotal, 3).Value = Cells(i, 14).Value
            wsTotal.Cells(filaTotal, 4).Value = 0
            wsTotal.Cells(filaTotal, 5).Value = 0
            wsTotal.Cells(filaTotal, 6).Value = 0
            wsTotal.Cells(filaTotal, 7).Value = Cells(i, 15).Value
            wsTotal.Cells(filaTotal, 8).Value = Cells(i, 23).Value
            wsTotal.Cells(filaTotal, 9).Value = Cells(i, 24).Value
            wsTotal.Cells(filaTotal, 10).Value = Cells(i, 25).Value
            ultDoc = Cells(i, 12).Value
            
            If Cells(i, 4).Value = 1 Then
                wsTotal.Cells(filaTotal, 4).Value = wsTotal.Cells(filaTotal, 4).Value + Cells(i, 7).Value
            Else
                If Cells(i, 4).Value = 100 Then
                    wsTotal.Cells(filaTotal, 5).Value = wsTotal.Cells(filaTotal, 5).Value + Cells(i, 7).Value
                Else
                    If Cells(i, 4).Value = 246 Then
                        wsTotal.Cells(filaTotal, 6).Value = wsTotal.Cells(filaTotal, 6).Value + Cells(i, 7).Value
                    End If
                End If
            End If
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Listar_Cobra_Cptos()
    Dim nFilas As Long
    Dim rango As Range
    Dim wsTotal As Excel.Worksheet
    Dim i As Long
    Dim ultDoc As String
    Dim nColumnas As Integer
    Dim ultCpto As Integer
    
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Resultado"
    Application.DisplayAlerts = True
    
    Set wsTotal = Worksheets("Resultado")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    MsgBox "Debe estar ordenado por DNI y Cpto.", , "Atención!!"
    
    'Encabezado Hoja Totales
    Range("1:1").Copy Destination:=wsTotal.Range("1:1")
    filaTotal = 2
    
    
    For i = 2 To nFilas - 1
        If Cells(i, 12).Value = Cells(i + 1, 12).Value Then
            Do
                Range(i & ":" & i).Copy Destination:=wsTotal.Range(filaTotal & ":" & filaTotal)
                i = i + 1
                filaTotal = filaTotal + 1
            Loop While Cells(i, 12).Value = Cells(i + 1, 12).Value
            Range(i & ":" & i).Copy Destination:=wsTotal.Range(filaTotal & ":" & filaTotal)
            filaTotal = filaTotal + 1
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Control_Duplicado()
    Dim nFilas As Long
    Dim rango As Range
    Dim wsTotal As Excel.Worksheet
    Dim i As Long
    Dim ultDoc As String
    Dim nColumnas As Integer
    Dim ultCpto As Integer
    
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Repetidos"
    Application.DisplayAlerts = True
    
    Set wsTotal = Worksheets("Repetidos")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    MsgBox "Debe estar ordenado por DNI.", , "Atención!!"
    
    'Encabezado Hoja Totales
    wsTotal.Cells(1, 1).Value = "Cuof"
    wsTotal.Cells(1, 2).Value = "DNI"
    wsTotal.Cells(1, 3).Value = "Nombre"
    wsTotal.Cells(1, 4).Value = "T.Prof"
    wsTotal.Cells(1, 5).Value = "Cpto"
    wsTotal.Cells(1, 6).Value = "Horas"
    filaTotal = 2
    
    
    For i = 2 To nFilas - 1
        bandera = True
        For j = 1 To 11
            If Cells(i, j).Value <> Cells(i + 1, j).Value Then
                bandera = False
            End If
        Next j
        
        If bandera Then
            Cells(i, nColumnas + 1).Value = "Repetido"
            Cells(i + 1, nColumnas + 1).Value = "Repetido"
            
            wsTotal.Cells(filaTotal, 1).Value = Cells(i, 1).Value
            wsTotal.Cells(filaTotal, 2).Value = Cells(i, 5).Value
            wsTotal.Cells(filaTotal, 3).Value = Cells(i, 6).Value
            wsTotal.Cells(filaTotal, 4).Value = Cells(i, 7).Value
            wsTotal.Cells(filaTotal, 5).Value = Cells(i, 8).Value
            If Cells(i, 9).Value <> "" Then
                wsTotal.Cells(filaTotal, 6).Value = Cells(i, 9).Value
            Else
                If Cells(i, 10).Value <> "" Then
                    wsTotal.Cells(filaTotal, 6).Value = Cells(i, 10).Value
                Else
                    wsTotal.Cells(filaTotal, 6).Value = Cells(i, 11).Value
                End If
            End If
            filaTotal = filaTotal + 1
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub


Sub Controlar_Doc()
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim i As Integer
    
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    For i = 2 To (nFilas - 1)
        bandera = True
        valorDoc = Cells(i, 3).Value
        For j = 1 To nColumnas
            If Cells(i, j).Value <> Cells(i + 1, j).Value Then
                bandera = False
            End If
        Next j
        If bandera Then
            Cells(i + 1, nColumnas + 1).Value = "Repetido"
        End If
    Next i
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

