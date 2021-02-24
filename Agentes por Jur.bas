Attribute VB_Name = "Módulo11"
Sub AgentesxJur()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i, cont As Long
    

    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    cont = 0
    For i = 2 To 52
      For J = 2 To nFilas
       If Cells(i, 39).Value = Cells(J, 5).Value Then
        cont = cont + 1
       End If
       Cells(i, 41).Value = cont
       cont = 0
      Next J
    Next i
   
    MsgBox "Proceso Exitoso"
End Sub


