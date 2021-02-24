Attribute VB_Name = "Módulo1"

Sub total_cuotas_por_agentes()
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
    dni = Cells(2, 5).Value
    monto = 0
    band = False
    
    For J = 2 To nFilas
    
      If dni = Cells(J, 5).Value Then
         
            monto = monto + Cells(J, 16).Value
            pos = J
           Else
            Cells(pos, 23).Value = monto
            monto = 0
            dni = Cells(J, 5).Value
            monto = monto + Cells(J, 16).Value
          
           
       End If
    Next J
    MsgBox "Proceso Exitoso"
End Sub
