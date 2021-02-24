Attribute VB_Name = "Módulo1"
Sub Cuotas()
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim K As Integer
    Dim J As Integer
    Dim cont, monto As Integer
    Dim band As Boolean
    
    'Debe estar ordenado por dni y por importe de mayor a menor
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    pos = 0
    dni = Cells(2, 4).Value
    monto = 0
    band = False
    For J = 2 To 35
    
      If dni = Cells(J, 4).Value Then
         If monto < 2000 Then
            monto = monto + Cells(2, 20).Value
            Cells(J, 22).Value = "42020"
            pos = J
            
           Else
            If band = False Then
             Cells(pos, 23).Value = monto
             monto = 0
             band = True
            End If
         End If
         
        Else
        
           
           dni = Cells(2, 4).Value
           
           band = False
           
           If monto < 2000 Then
                monto = monto + Cells(2, 20).Value
                Cells(J, 22).Value = "42020"
                pos = J
             Else
                If band = False Then
                 Cells(pos, 23).Value = monto
                 monto = 0
                 band = True
                End If
            End If
         
        End If
        
      Next J

    MsgBox "Proceso Finalizado"
End Sub


