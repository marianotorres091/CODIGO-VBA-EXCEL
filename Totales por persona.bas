Attribute VB_Name = "Módulo1"
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
    
    ultDoc = Cells(2, 5).Value
    importe = 0
    ultJur = Cells(2, 2).Value
    nombre = Cells(2, 7).Value
    limite = nFilas
    
    For i = 2 To limite
    Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
        If Cells(i, 8).Value < 400 Then
            If ultDoc = Cells(i, 5).Value Then
                If Cells(i, 9).Value = 2 Then
                    importe = importe - Cells(i, 11).Value
                Else
                    importe = importe + Cells(i, 11).Value
                End If
            Else
                wsTotal.Cells(filaTotal, 1).Value = ultJur
                wsTotal.Cells(filaTotal, 2).Value = ultDoc
                wsTotal.Cells(filaTotal, 3).Value = nombre
                wsTotal.Cells(filaTotal, 4).Value = importe
                filaTotal = filaTotal + 1
                
                ultDoc = Cells(i, 5).Value
                importe = 0
                ultJur = Cells(i, 2).Value
                nombre = Cells(i, 7).Value
                i = i - 1
            End If
        Else
           If ultDoc = Cells(i, 5).Value Then
                If Cells(i, 9).Value = 1 Then
                    importe = importe - Cells(i, 11).Value
                Else
                    importe = importe + Cells(i, 11).Value
                End If
            Else
                wsTotal.Cells(filaTotal, 1).Value = ultJur
                wsTotal.Cells(filaTotal, 2).Value = ultDoc
                wsTotal.Cells(filaTotal, 3).Value = nombre
                wsTotal.Cells(filaTotal, 4).Value = importe
                filaTotal = filaTotal + 1
                
                ultDoc = Cells(i, 5).Value
                importe = 0
                ultJur = Cells(i, 2).Value
                nombre = Cells(i, 7).Value
                i = i - 1
            End If
        End If
    Next i
    wsTotal.Cells(filaTotal, 1).Value = ultJur
    wsTotal.Cells(filaTotal, 2).Value = ultDoc
    wsTotal.Cells(filaTotal, 3).Value = nombre
    wsTotal.Cells(filaTotal, 4).Value = importe
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Application.StatusBar = False
End Sub

