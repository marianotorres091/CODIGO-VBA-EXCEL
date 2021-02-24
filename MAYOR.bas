Attribute VB_Name = "Módulo1"
Sub MAYOR()
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim K As Integer
    Dim J As Integer
    Dim cont As Integer
    Dim band As Boolean
    
    'Debe estar ordenado por dni y por importe de mayor a menor
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
   x = 2
    pos = Cells(2, 8).Value
    monto = Cells(2, 14).Value
    band = False
    For J = 2 To nFilas
    
      If pos = Cells(J, 8).Value Then
         If Cells(J, 14).Value > monto Then
            Cells(J, 9).Value = "MAYOR"
            monto = Cells(J, 14).Value
            
           Else
          
            If Cells(J, 14).Value < monto Then
             Cells(x, 9).Value = "MAYOR"
             monto = Cells(x, 14).Value
           
           
             Else
             
               If Cells(J, 14).Value = monto Then
                  If band = False Then
                    Cells(x, 9).Value = "MAYOR"
                    monto = Cells(J, 14).Value
                    band = True
                  End If
               End If
            End If
           
         End If
      End If
     
        If pos <> Cells(J, 8).Value Then
         pos = Cells(J, 8).Value
         monto = Cells(J, 14).Value
         x = J
         band = False
        End If
        
    Next J

    MsgBox "Proceso Finalizado"
End Sub
