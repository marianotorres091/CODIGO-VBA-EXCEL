Attribute VB_Name = "Módulo11"
Sub Totales_Jur()
    Dim total_mov As Long
    Dim total_dni As Long
    Dim total_actuaciones As Long
    Dim nFilas As Long
    Dim filaTotal As Long
    Dim rango As Range
    Dim wsTotal As Excel.Worksheet
    Dim i As Long
    Dim ultDoc As String
    Dim ultJur As Integer
    
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Totales"
    Application.DisplayAlerts = True
    
    Set wsTotal = Worksheets("Totales")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    'Encabezado Hoja Totales
    wsTotal.Cells(1, 1).Value = "Jurisdicción"
    wsTotal.Cells(1, 2).Value = "Total de Documentos"
    wsTotal.Cells(1, 3).Value = "Total de Movimientos"
    wsTotal.Cells(1, 4).Value = "Total de Actuaciones"
    filaTotal = 2
    
    ultDoc = Cells(2, 5).Value
    ultJur = Cells(2, 2).Value
    total_actuaciones = 0
    total_dni = 1
    total_mov = 1
    
    For i = 3 To nFilas
        If ultJur = Cells(i, 2).Value Then
            If ultDoc = Cells(i, 5).Value Then
                total_mov = total_mov + 1
            Else
                ultDoc = Cells(i, 5).Value
                total_dni = total_dni + 1
                total_mov = total_mov + 1
            End If
        Else
            wsTotal.Cells(filaTotal, 1).Value = ultJur
            wsTotal.Cells(filaTotal, 2).Value = total_dni
            wsTotal.Cells(filaTotal, 3).Value = total_mov
            wsTotal.Cells(filaTotal, 4).Value = total_actuaciones
            filaTotal = filaTotal + 1
            ultJur = Cells(i, 2).Value
            ultDoc = Cells(i, 5).Value
            total_actuaciones = 0
            total_dni = 1
            total_mov = 1
        End If
    Next i
    wsTotal.Cells(filaTotal, 1).Value = ultJur
    wsTotal.Cells(filaTotal, 2).Value = total_dni
    wsTotal.Cells(filaTotal, 3).Value = total_mov
    wsTotal.Cells(filaTotal, 4).Value = total_actuaciones
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Totales_Cpto()
    Dim total_mov As Long
    Dim total_dni As Long
    Dim total_actuaciones As Long
    Dim nFilas As Long
    Dim filaTotal As Long
    Dim rango As Range
    Dim wsTotal As Excel.Worksheet
    Dim i As Long
    Dim ultDoc As String
    Dim ultJur As Integer
    
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Totales"
    Application.DisplayAlerts = True
    
    Set wsTotal = Worksheets("Totales")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    'Encabezado Hoja Totales
    wsTotal.Cells(1, 1).Value = "Documento"
    wsTotal.Cells(1, 2).Value = "CPTO"
    wsTotal.Cells(1, 3).Value = "Importe"
    filaTotal = 2
    
    ultDoc = Cells(2, 12).Value
    ultCpto = Cells(2, 4).Value
    importe = 0
    
    For i = 2 To nFilas
        If ultCpto = Cells(i, 4).Value And ultDoc = Cells(i, 12).Value Then
            importe = importe + Cells(i, 7).Value
        Else
            wsTotal.Cells(filaTotal, 1).Value = ultDoc
            wsTotal.Cells(filaTotal, 2).Value = ultCpto
            wsTotal.Cells(filaTotal, 3).Value = importe
            filaTotal = filaTotal + 1
            ultCpto = Cells(i, 4).Value
            ultDoc = Cells(i, 12).Value
            importe = Cells(i, 7).Value
        End If
    Next i
    wsTotal.Cells(filaTotal, 1).Value = ultDoc
    wsTotal.Cells(filaTotal, 2).Value = ultCpto
    wsTotal.Cells(filaTotal, 3).Value = importe
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Promedio_Cpto()
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i As Long
    Dim ultDoc As String
    Dim ultJur As Integer
    
    
    'Regresa el control a la hoja de origen
    Sheets("Totales").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    ultCpto = Cells(2, 2).Value
    importe = 0
    contador = 0
    fila = 2
    
    For i = 2 To nFilas
        If ultCpto = Cells(i, 2).Value Then
            importe = importe + Cells(i, 3).Value
            contador = contador + 1
        Else
            Cells(fila, nColumnas + 4).Value = ultCpto
            Cells(fila, nColumnas + 5).Value = importe / contador
            fila = fila + 1
            ultCpto = Cells(i, 2).Value
            importe = Cells(i, 3).Value
            contador = 1
        End If
    Next i
    Cells(fila, nColumnas + 4).Value = ultCpto
    Cells(fila, nColumnas + 5).Value = importe / contador
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Totales_Persona()
    Dim total_mov As Long
    Dim total_dni As Long
    Dim total_actuaciones As Long
    Dim nFilas As Long
    Dim filaTotal As Long
    Dim rango As Range
    Dim wsTotal As Excel.Worksheet
    Dim i As Long
    Dim ultDoc As String
    Dim ultJur As Integer
    
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Total x Persona"
    Application.DisplayAlerts = True
    
    Set wsTotal = Worksheets("Total x Persona")
    
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
    wsTotal.Cells(1, 4).Value = "Importe"
    filaTotal = 2
    
    ultDoc = Cells(2, 12).Value
    importe = 0
    ultJur = Cells(2, 8).Value
    nombre = Cells(2, 5).Value
    
    For i = 2 To nFilas
        If Cells(i, 4).Value < 350 Then
            If ultDoc = Cells(i, 12).Value Then
                If Cells(i, 6).Value = 2 Then
                    importe = importe - Cells(i, 7).Value
                Else
                    importe = importe + Cells(i, 7).Value
                End If
            Else
                wsTotal.Cells(filaTotal, 1).Value = ultJur
                wsTotal.Cells(filaTotal, 2).Value = ultDoc
                wsTotal.Cells(filaTotal, 3).Value = nombre
                wsTotal.Cells(filaTotal, 4).Value = importe
                filaTotal = filaTotal + 1
                
                ultDoc = Cells(i, 12).Value
                importe = 0
                ultJur = Cells(i, 8).Value
                nombre = Cells(i, 5).Value
                i = i - 1
            End If
        End If
    Next i
    wsTotal.Cells(filaTotal, 1).Value = ultJur
    wsTotal.Cells(filaTotal, 2).Value = ultDoc
    wsTotal.Cells(filaTotal, 3).Value = nombre
    wsTotal.Cells(filaTotal, 4).Value = importe
    
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
    
    For i = 2 To nFilas
        If Cells(i, 4).Value < 350 Then
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
                filaTotal = filaTotal + 1
                ultCpto = Cells(i, 4).Value
                ultJur = Cells(i, 8).Value
                importe = 0
                i = i - 1
            End If
        End If
    Next i
    wsTotal.Cells(filaTotal, 1).Value = ultJur
    wsTotal.Cells(filaTotal, 2).Value = ultCpto
    wsTotal.Cells(filaTotal, 3).Value = descripcion
    wsTotal.Cells(filaTotal, 4).Value = importe
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

