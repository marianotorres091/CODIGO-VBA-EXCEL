Attribute VB_Name = "M�dulo11"
Sub Subrogancia()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i, cont, mes, a�o, mesa�o As Long
    

    'Calcular el n�mero de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    cont2 = 1
    cont = 1
    mes = 10000
    a�o = 2014
    mesa�o = mes + a�o
     cont = cont + 1
     
      For j = 1 To 72
       Cells(cont, 9).Value = "1"
       Cells(cont, 12).Value = mesa�o
       Cells(cont + 1, 9).Value = "7"
       Cells(cont + 1, 12).Value = mesa�o
       Cells(cont + 2, 9).Value = "16"
       Cells(cont + 2, 12).Value = mesa�o
       Cells(cont + 3, 9).Value = "25"
       Cells(cont + 3, 12).Value = mesa�o
       Cells(cont + 4, 9).Value = "126"
       Cells(cont + 4, 12).Value = mesa�o
       Cells(cont + 5, 9).Value = "154"
       Cells(cont + 5, 12).Value = mesa�o
       Cells(cont + 6, 9).Value = "200"
       Cells(cont + 6, 12).Value = mesa�o
       Cells(cont + 7, 9).Value = "213"
       Cells(cont + 7, 12).Value = mesa�o
       cont = cont + 8
       
       cont2 = cont2 + 1
       mes = mes + 10000
       mesa�o = mes + a�o
       
       If cont2 = 12 Then
       a�o = a�o + 1
       mes = 0
       cont2 = 0
       End If
       
      Next j
   
    MsgBox "Proceso Exitoso"
End Sub


