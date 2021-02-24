Attribute VB_Name = "Módulo11"
Sub Subrogancia()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i, cont, mes, año, mesaño As Long
    

    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    cont2 = 1
    cont = 1
    mes = 10000
    año = 2014
    mesaño = mes + año
     cont = cont + 1
     
      For j = 1 To 72
       Cells(cont, 9).Value = "1"
       Cells(cont, 12).Value = mesaño
       Cells(cont + 1, 9).Value = "7"
       Cells(cont + 1, 12).Value = mesaño
       Cells(cont + 2, 9).Value = "16"
       Cells(cont + 2, 12).Value = mesaño
       Cells(cont + 3, 9).Value = "25"
       Cells(cont + 3, 12).Value = mesaño
       Cells(cont + 4, 9).Value = "126"
       Cells(cont + 4, 12).Value = mesaño
       Cells(cont + 5, 9).Value = "154"
       Cells(cont + 5, 12).Value = mesaño
       Cells(cont + 6, 9).Value = "200"
       Cells(cont + 6, 12).Value = mesaño
       Cells(cont + 7, 9).Value = "213"
       Cells(cont + 7, 12).Value = mesaño
       cont = cont + 8
       
       cont2 = cont2 + 1
       mes = mes + 10000
       mesaño = mes + año
       
       If cont2 = 12 Then
       año = año + 1
       mes = 0
       cont2 = 0
       End If
       
      Next j
   
    MsgBox "Proceso Exitoso"
End Sub


