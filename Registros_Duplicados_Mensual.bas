Attribute VB_Name = "Módulo1"
Sub Registros_duplicados_Mensual()
    Dim nFilas As Long
    Dim nColumnas As Long
    'Dim i As Integer
    Dim rango As Range
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
   
    limite = (nFilas - 1)
    For i = 2 To limite
     Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
     If Cells(i, nColumnas + 1).Value <> "Repetido" Then
     
        'valorjur = Cells(i, 2).Value
        'valorescalafon = Cells(i, 3).Value
        valorDoc = Cells(i, 5).Value
        valorCouc = Cells(i, 8).Value
        valorRj = Cells(i, 10).Value
        valorunidad = Cells(i, 11).Value
        valorimporte = Cells(i, 12).Value
        valorVto = Cells(i, 14).Value
        'cuota = Cells(i, 13).Value
        band = False
        
        For j = (i + 1) To nFilas
          If valorDoc = Cells(j, 5).Value Then
             band = True
            If valorDoc = Cells(j, 5).Value And valorCouc = Cells(j, 8).Value And valorRj = Cells(j, 10).Value And valorunidad = Cells(j, 11).Value And valorimporte = Cells(j, 12).Value And valorVto = Cells(j, 14).Value Then
                
                Cells(j, 5).Interior.Color = RGB(153, 196, 195)
                Cells(i, 5).Interior.Color = RGB(153, 196, 195)
                Cells(j, nColumnas + 1).Value = "Repetido"
                
                If Cells(j, nColumnas + 2).Value = "" Then
                  Cells(j, nColumnas + 2).Value = i
                  Cells(i, nColumnas + 2).Value = i
                Else
                  Cells(i, nColumnas + 2).Value = Cells(j, nColumnas + 2).Value
                End If
                
                Cells(i, nColumnas + 1).Value = "Repetido"
            End If
          Else
            If band Then
              j = nFilas
            End If
          End If
        Next j
      End If
    Next i
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Application.StatusBar = False
End Sub

   





