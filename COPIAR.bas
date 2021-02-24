Attribute VB_Name = "Módulo1"
Sub Comparar_mismo_archivo()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i, cuoc, rj, unidad, importe, vto As Long
   

    
    limite = 15
    For i = 3 To limite
       
        'Cells(i, 35).Value = "buscado"
        Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
       
      
      For j = 2 To 19
        
       
         If Cells(i, 20).Value = Cells(j, 8).Value Then
            If Cells(j, 18).Value = 122019 Then
                 If Cells(j, 10).Value = 0 Then
                   Cells(i, 22).Value = Cells(j, 12).Value
                   Else
                    Cells(i, 23).Value = Cells(j, 12).Value
                 End If
              Else
                  If Cells(j, 18).Value = 12020 Then
                    If Cells(j, 10).Value = 0 Then
                      Cells(i, 24).Value = Cells(j, 12).Value
                      Else
                       Cells(i, 25).Value = Cells(j, 12).Value
                    End If
                   Else
                       If Cells(j, 18).Value = 22020 Then
                            If Cells(j, 10).Value = 0 Then
                              Cells(i, 26).Value = Cells(j, 12).Value
                              Else
                               Cells(i, 27).Value = Cells(j, 12).Value
                            End If
                        Else
                            If Cells(j, 18).Value = 32020 Then
                                If Cells(j, 10).Value = 0 Then
                                  Cells(i, 28).Value = Cells(j, 12).Value
                                  Else
                                   Cells(i, 29).Value = Cells(j, 12).Value
                                End If
                             Else
                                 If Cells(j, 18).Value = 42020 Then
                                    If Cells(j, 10).Value = 0 Then
                                      Cells(i, 30).Value = Cells(j, 12).Value
                                      Else
                                       Cells(i, 31).Value = Cells(j, 12).Value
                                    End If
                                   Else
                                       If Cells(j, 18).Value = 52020 Then
                                            If Cells(j, 10).Value = 0 Then
                                              Cells(i, 32).Value = Cells(j, 12).Value
                                              Else
                                               Cells(i, 33).Value = Cells(j, 12).Value
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

