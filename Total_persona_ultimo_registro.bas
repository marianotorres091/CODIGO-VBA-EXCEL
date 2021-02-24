Attribute VB_Name = "Módulo11"
Sub Totales_Persona_ultimo_registro()
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
    
  
    
    'Regresa el control a la hoja de origen
    Sheets("VER DE WR - Descuento Cuotas").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    MsgBox "Debe estar ordenado por DNI.", , "Atención!!"
    
   
    
    ultDoc = Cells(2, 5).Value
    importe = 0
    ultJur = Cells(2, 2).Value
    nombre = Cells(2, 7).Value
    
    For i = 2 To nFilas
        If Cells(i, 4).Value < 350 Then
            If ultDoc = Cells(i, 5).Value Then
                If Cells(i, 9).Value = 2 Then
                    importe = importe - Cells(i, 11).Value
                Else
                    importe = importe + Cells(i, 11).Value
                End If
            Else
                
               Cells(i - 1, 15).Value = importe
               
                
                ultDoc = Cells(i, 5).Value
                importe = 0
                ultJur = Cells(i, 2).Value
                nombre = Cells(i, 7).Value
                i = i - 1
            End If
        End If
    Next i
   
     Cells(i - 1, 15).Value = importe
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub



