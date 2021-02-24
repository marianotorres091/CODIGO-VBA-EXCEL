Attribute VB_Name = "Módulo1"
Sub Comparar_mismo_archivo()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i, cuoc, rj, unidad, importe, vto As Long
   

    
    limite = 3668
    For i = 2 To limite
       
        'Cells(i, 35).Value = "buscado"
        Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
       
      
      For j = 2 To 6622
        
       
         If Cells(i, 34).Value = Cells(j, 1).Value Then
            If Cells(i, 35).Value = Cells(j, 2).Value Then
                 If Cells(i, 37).Value = Cells(j, 4).Value Then
                    If Cells(i, 40).Value = Cells(j, 7).Value Then
                        If Cells(i, 41).Value = Cells(j, 8).Value Then
                          If Cells(i, 42).Value = Cells(j, 9).Value Then
                             If Cells(i, 43).Value = Cells(j, 10).Value Then
                                 If Cells(i, 44).Value = Cells(j, 11).Value Then
                                    If Cells(i, 45).Value = Cells(j, 12).Value Then
                                      If Cells(i, 46).Value = Cells(j, 13).Value Then
                                         'If Cells(i, 49).Value = Cells(j, 16).Value Then
                                            
                                   
                                            Cells(j, 27).Value = Cells(i, 52).Value
                                            Cells(j, 28).Value = Cells(i, 53).Value
                                            Cells(j, 29).Value = Cells(i, 54).Value
                                            Cells(j, 30).Value = Cells(i, 55).Value
                                            Cells(j, 31).Value = Cells(i, 56).Value
                                            Cells(j, 32).Value = Cells(i, 57).Value
                                            Cells(j, 33).Value = Cells(i, 58).Value
                                            
                                        ' End If
                                      End If
                                    End If
                                 End If
                             End If
                          End If
                        End If
                     End If
                  End If
            End If
         End If
          
          
      Next j
       
    Next i
     MsgBox "Proceso exitosa"
      Application.StatusBar = False
End Sub

