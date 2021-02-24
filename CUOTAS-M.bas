Attribute VB_Name = "Módulo1"
Sub Cuotas_M()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i, cont As Long
    

    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    band = False
    monto = 0
    For i = 2 To 4
     
      For J = 2 To 8
       
           If Cells(J, 5).Value = Cells(i, 29).Value Then
                   
                If monto < 30000 Then
                      monto = monto + Cells(J, 16).Value
                     If monto < 30000 Then
                      montofinal = monto
                      Cells(J, 21).Value = "4"
                      Cells(J, 22).Value = "2020"
                      pos = J
                     End If
                   Else
                      If band = False Then
                       Cells(pos, 23).Value = montofinal
                       band = True
                      End If
                 End If
             Else
             If Cells(pos, 23).Value = "" Then
              If band = False Then
                 Cells(pos, 23).Value = montofinal
                 monto = 0
                 band = True
                 
              End If
             End If
            End If
           
      
      Next J
      band = False
    Next i
   
    MsgBox "Proceso Exitoso"
End Sub
