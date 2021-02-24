Attribute VB_Name = "Módulo11"
Sub Cuotas_DESCUENTOS_PERSONA()
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
    Dim band As Double
    
  
    
    'Regresa el control a la hoja de origen
    Sheets("VER DE WR - Descuento Cuotas").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
 
    'usando como porcentaje de importe a comparar un valor negativo osea el desc tiene que ser negativo
   
    band = False
    ultDoc = Cells(2, 5).Value
    importe = 0
    ultJur = Cells(2, 2).Value
    nombre = Cells(2, 7).Value
    desc = 0
    For i = 2 To nFilas
        If Cells(i, 4).Value < 350 Then
          For j = 2 To 45
          If ultDoc = Cells(j, 18).Value Then
           desc = Cells(j, 21).Value
          End If
          Next j
            If ultDoc = Cells(i, 5).Value Then
               If Cells(i, 9).Value = 2 Then
                    If importe > desc Then
                            importe = importe - Cells(i, 11).Value
                         If importe >= desc Then
                            importefinal = importe
                            Cells(i, 16).Value = importefinal
                          Else
                            importe = importe + Cells(i, 11).Value
                            If importe >= desc Then
                             If band = False Then
                                importefinal = importe
                                Cells(i - 1, 17).Value = "cuota1"
                                importe = importe - Cells(i, 11).Value
                                band = True
                             End If
                            End If
                         End If
                       Else
                        If i = 2 Then
                         importe = importe - Cells(i, 11).Value
                        End If
                     End If
                Else
                    importe = importe + Cells(i, 11).Value
                End If
            Else
              
              If importe >= desc Then
                 If band = False Then
                    importefinal = importe
                    Cells(i - 1, 17).Value = "cuota1"
                    band = True
                 End If
              End If
               band = False
                
                ultDoc = Cells(i, 5).Value
                importe = 0
                ultJur = Cells(i, 2).Value
                nombre = Cells(i, 7).Value
                i = i - 1
            End If
        End If
    Next i
   If importe >= desc Then
    importefinal = importe
    Cells(i - 1, 17).Value = "cuota1"
   End If
     
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub



