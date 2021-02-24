Attribute VB_Name = "Módulo1"
Sub Comparar_mismo_archivo()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i, cuoc, rj, unidad, importe, vto As Long
    jur = 3
    ceic = 1
    cargo = 2
    limite = 31
    cantidad = 0
    
  For jur = 3 To limite
        'Cells(i, 35).Value = "buscado"
        Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
       
      
   
     
     For i = 2 To 455
        For fila = 2 To 37
       
         If Cells(i, 35).Value = Cells(1, jur).Value Then
            If Cells(i, 37).Value = Cells(fila, 1).Value Then
                 If Cells(i, 39).Value = Cells(fila, 2).Value Then
                    If Cells(i, 42).Value = Cells(fila, jur).Value Then
                        
                                    Cells(i, 43).Value = "son iguales"
                            
                    End If
                  End If
            End If
         End If
       Next fila
     Next i
  Next jur
       
    
     MsgBox "Proceso exitosa"
      Application.StatusBar = False
End Sub

