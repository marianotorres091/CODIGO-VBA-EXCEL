Attribute VB_Name = "Módulo1"
Sub Carga()
    Dim wsTabla As Worksheet, _
        wsDestino As Worksheet
    Dim nFilas As Long
    Dim i As Long
    
    ThisWorkbook.Activate
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Resultados"
    Application.DisplayAlerts = True
    
    Set wsDestino = Worksheets("Resultados")
    Set wsTabla = Worksheets("Hoja2")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select

    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    'Fila del encabezado de Destino
    wsDestino.Cells(1, 1).Value = "PtaId"
    wsDestino.Cells(1, 2).Value = "JurId"
    wsDestino.Cells(1, 3).Value = "EscId"
    wsDestino.Cells(1, 4).Value = "Pref"
    wsDestino.Cells(1, 5).Value = "Doc"
    wsDestino.Cells(1, 6).Value = "Digito"
    wsDestino.Cells(1, 7).Value = "Nombres"
    wsDestino.Cells(1, 8).Value = "Couc"
    wsDestino.Cells(1, 9).Value = "Reajuste"
    wsDestino.Cells(1, 10).Value = "Unidades"
    wsDestino.Cells(1, 11).Value = "Importe"
    wsDestino.Cells(1, 12).Value = "Vto"
    wsDestino.Range("1:1").Font.Bold = True
    wsDestino.Range("1:1").HorizontalAlignment = xlCenter
    
    nFilasDest = 2
    
    
    For i = 2 To nFilas
        For j = 2 To 15
            wsDestino.Cells(nFilasDest, 1).Value = 0
            wsDestino.Cells(nFilasDest, 2).Value = 36
            wsDestino.Cells(nFilasDest, 3).Value = 2
            wsDestino.Cells(nFilasDest, 4).Value = 0
            wsDestino.Cells(nFilasDest, 5).Value = Cells(i, 3).Value
            wsDestino.Cells(nFilasDest, 6).Value = 0
            wsDestino.Cells(nFilasDest, 7).Value = Cells(i, 6).Value
            wsDestino.Cells(nFilasDest, 8).Value = 212
            wsDestino.Cells(nFilasDest, 9).Value = 1
            wsDestino.Cells(nFilasDest, 10).Value = 0
            wsDestino.Cells(nFilasDest, 11).Value = wsTabla.Cells(j, 3).Value
            wsDestino.Cells(nFilasDest, 12).Value = wsTabla.Cells(j, 4).Value
            
            nFilasDest = nFilasDest + 1
        Next j
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

