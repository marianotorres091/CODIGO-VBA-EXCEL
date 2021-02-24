Attribute VB_Name = "Módulo1"
Sub BuscarDuplicados()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i As Long
    

    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count

    For i = 2 To nFilas
       Cells(i, 16).Value = "buscado"
       Cells(i, 18).Value = i
      For j = 2 To nFilas
        If Cells(j, 16).Value <> "buscado" Then
            If Cells(i, 11).Value = Cells(j, 11).Value Then
              If Cells(i, 12).Value = Cells(j, 12).Value Then
              Cells(i, 17).Value = "repetido"
              Cells(j, 19).Value = i
              End If
            End If
        End If
      Next j
      Cells(i, 16).Value = " "
    Next i
   
    MsgBox "Proceso Exitoso"
End Sub
