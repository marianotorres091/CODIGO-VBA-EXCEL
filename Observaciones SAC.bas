Attribute VB_Name = "Módulo1"
Sub OBSERVACIONES_SAC()

    Dim nFilas As Double
    Dim nColumnas As Double
    Dim K As Integer
    Dim J As Integer
    Dim cont As Integer
    Dim band As Boolean
    
    'CALCULO DEL PRIMER Y SEGUNDO MAYOR SAC (TIENEN 6 COLUMNAS DE ACUMULADOS POR AGENTE)
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
  
    mayor1 = Cells(2, 12).Value
    posigual = 0
    J = 12
    band = False
    bandigual = False
    bandasigmayor2 = False
    
    limite = 133367
    
    For i = 2 To limite
       
       Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
      
       mayor1 = Cells(i, 12).Value
       posmayor1 = 12
       posigual = 0
       band = False
       bandigual = False
       contcero = 0
       
        For t = 12 To 16
         If Cells(i, t).Value <> 0 Then
            If Cells(i, t).Value = mayor1 Then
              bandigual = True
             Else
             bandigual = False
            End If
           Else
           contcero = contcero + 1
         End If
        Next t
       
        If contcero = 4 Then
          Cells(i, 27).Value = "tiene 1 acumulados"
          Else
            If contcero = 3 Then
              Cells(i, 27).Value = "tiene 2 acumulados"
               Else
                If contcero = 2 Then
                  Cells(i, 27).Value = "tiene 3 acumulados"
                  Else
                   If contcero = 1 Then
                    Cells(i, 27).Value = "tiene 4 acumulados"
                    Else
                      If contcero = 0 Then
                      Cells(i, 27).Value = "tiene 5 acumulados"
                      End If
                   End If
                End If
            End If
        End If
         
         If bandigual = True Then
          Cells(i, 28).Value = "todos iguales"
         End If
    Next i
   MsgBox "Proceso exitoso"
   Application.StatusBar = False
End Sub
         








