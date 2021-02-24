Attribute VB_Name = "Módulo1"
Sub Totales_Registros_Persona()
    Dim total_mov As Long
    Dim total_dni As Long
    Dim total_actuaciones As Long
    Dim nFilas As Long
    Dim filaTotal As Long
    Dim rango As Range
    Dim wsTotal As Excel.Worksheet
    Dim wsHojaVieja As Excel.Worksheet
    Dim i As Long
    Dim ultDoc As String
    Dim ultJur As Integer
    
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Total Filas x Persona"
    Application.DisplayAlerts = True
    
    Set wsTotal = Worksheets("Total Filas x Persona")
    
    'Regresa el control a la hoja de origen
    Sheets("DESCUENTOS-HISTORICO").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    MsgBox "Debe estar ordenado por DNI.", , "Atención!!"
    
    'Encabezado Hoja Totales
    wsTotal.Cells(1, 1).Value = "JUR"
    wsTotal.Cells(1, 2).Value = "DNI"
    wsTotal.Cells(1, 3).Value = "Nombre"
    wsTotal.Cells(1, 4).Value = "Nº Filas"
    filaTotal = 2
    
    ultDoc = Cells(2, 5).Value
    importe = 0
    ultJur = Cells(2, 2).Value
    nombre = Cells(2, 7).Value
    
    For i = 2 To nFilas
      If ultDoc = Cells(i, 5) Then
         cantfilas = cantfilas + 1
       Else
         Sheets("Total Filas x Persona").Select
             wsTotal.Cells(filaTotal, 1).Value = ultJur
             wsTotal.Cells(filaTotal, 2).Value = ultDoc
             wsTotal.Cells(filaTotal, 3).Value = nombre
             wsTotal.Cells(filaTotal, 4).Value = cantfilas
             filaTotal = filaTotal + 1
         Sheets("DESCUENTOS-HISTORICO").Select
              ultDoc = Sheets("DESCUENTOS-HISTORICO").Cells(i, 5).Value
              ultJur = Sheets("DESCUENTOS-HISTORICO").Cells(i, 2).Value
              nombre = Sheets("DESCUENTOS-HISTORICO").Cells(i, 7).Value
              cantfilas = 0
              cantfilas = cantfilas + 1
      End If
    Next i
             
             wsTotal.Cells(filaTotal, 1).Value = ultJur
             wsTotal.Cells(filaTotal, 2).Value = ultDoc
             wsTotal.Cells(filaTotal, 3).Value = nombre
             wsTotal.Cells(filaTotal, 4).Value = cantfilas
    
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub
