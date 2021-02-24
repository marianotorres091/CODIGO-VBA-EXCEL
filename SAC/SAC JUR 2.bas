Attribute VB_Name = "Módulo1"
Sub MAYORSAC_PASO1()
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
    
    Cells(1, 18).Value = "1º MAYOR"
    Cells(1, 19).Value = "POS 1º"
    Cells(1, 20).Value = "2º MAYOR"
    Cells(1, 21).Value = "POS 2º"
    Cells(1, 22).Value = "DIF % MAYOR1-MAYOR2"
    Cells(1, 23).Value = "ACUM SAC"
    Cells(1, 24).Value = "POSIBLE SAC"
    Cells(1, 25).Value = "OBSERVACIONES"
    
    mayor1 = Cells(2, 12).Value
    posigual = 0
    J = 12
    band = False
    bandigual = False
    bandasigmayor2 = False
    
    limite = nFilas
    
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
                If t = 12 Then
                 bandigual = True
                 Else
                  If bandigual Then
                    bandigual = True
                  End If
                End If
             Else
             bandigual = False
            End If
           Else
           contcero = contcero + 1
         End If
        Next t
       
        If contcero = 4 Then
          Cells(i, 25).Value = "tiene 1 acumulados"
          Else
           If contcero = 3 Then
             Cells(i, 25).Value = "tiene 2 acumulados"
             Else
              If contcero = 2 Then
                 Cells(i, 25).Value = "tiene 3 acumulados"
                 Else
                   If contcero = 1 Then
                      Cells(i, 25).Value = "tiene 4 acumulados"
                      Else
                       If contcero = 5 Then
                         Cells(i, 25).Value = "tiene 0 acumulados"
                         Else
                         If contcero = 0 Then
                            Cells(i, 25).Value = "tiene 5 acumulados"
                         End If
                       End If
                   End If
              End If
           End If
        End If
        
       'SI NO TIENE LOS 5 ACUM, CALCULO EL ACUM TOTAL Y OBTENGO EL SAC
        If contcero >= 1 Then
          acum = 0
          For col = 12 To 16
           acum = acum + Cells(i, col).Value
          Next col
        
        
            'GUARDO EL ACUM
            Cells(i, 23).Value = acum
            
            'CALCULO EL SAC
            sac = 0
            sac = acum / 12
            Cells(i, 24).Value = sac
            
        End If
        
     If contcero = 0 Then
      'CALCULO DEL PRIMER MAYOR
        If bandigual = False Then
        
            For J = 12 To 16
            
              
                 If Cells(i, J).Value > mayor1 Then
                    mayor1 = Cells(i, J).Value
                    Cells(i, 18).Value = mayor1
                    posmayor1 = J
                    
                   Else
                  
                    If Cells(i, J).Value < mayor1 Then
                    
                   
                   
                     Else
                     
                       If Cells(i, J).Value = mayor1 Then
                         If J <> 12 Then
                          posigual = J
                         End If
                       End If
                    End If
                   
                 End If

            Next J
              
             Cells(i, 18).Value = mayor1
             Cells(i, 19).Value = posmayor1
             
             Select Case posmayor1
              Case Is = 12
               Cells(i, 19).Value = "JUL"
              Case Is = 13
               Cells(i, 19).Value = "AGOS"
              Case Is = 14
               Cells(i, 19).Value = "SEPT"
              Case Is = 15
               Cells(i, 19).Value = "OCT"
              Case Is = 16
               Cells(i, 19).Value = "NOV"
             End Select
             
           'CALCULO DE 2DO MAYOR
         
            For h = 12 To 16
                If Cells(i, h).Value <> mayor1 Then
                 If bandasigmayor2 = False Then
                  mayor2 = Cells(i, h).Value
                  bandasigmayor2 = True
                  posmayor2 = h
                 End If
                End If
            Next h
             bandasigmayor2 = False
             
            For J = 12 To 16
            
              If Cells(i, J).Value <> mayor1 Then
              
                 If Cells(i, J).Value > mayor2 Then
                    mayor2 = Cells(i, J).Value
                    Cells(i, 20).Value = mayor2
                    posmayor2 = J
                    
                   Else
                  
                    If Cells(i, J).Value < mayor2 Then
                    
                   
                   
                     Else
                     
                       If Cells(i, J).Value = mayor2 Then
                         If J <> posmayor2 Then
                          posigual = J
                         End If
                       End If
                    End If
                   
                 End If
              End If
            Next J
             Cells(i, 20).Value = mayor2
             Cells(i, 21).Value = posmayor2
             
             Select Case posmayor2
              Case Is = 12
               Cells(i, 21).Value = "JUL"
              Case Is = 13
               Cells(i, 21).Value = "AGOS"
              Case Is = 14
               Cells(i, 21).Value = "SEPT"
              Case Is = 15
               Cells(i, 21).Value = "OCT"
              Case Is = 16
               Cells(i, 21).Value = "NOV"
             End Select
             
             
             
         Else
            If contcero = 0 Then
             Cells(i, 25).Value = "todos iguales"
             
             'CALCULO EL POSIBLE SAC
              saciguales = 0
              saciguales = mayor1 / 2
              Cells(i, 24).Value = saciguales
               
            End If
        End If
     End If
 Next i
    MsgBox "Proceso Finalizado"
    Application.StatusBar = False
End Sub


Sub SAC_Y_DIF_MAYOR1_MAYOR2_PASO2()

Dim nFilas As Double
Dim nColumnas As Double

'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count

'SE CALCULA LA DIFERENCIA QUE EXISTE ENTRE EL 1ER MAYOR Y EL 2DO MAYOR

 
 limite = nFilas
 
  For i = 2 To limite
  dif = 0
  sac = 0
  difporcentual = 0
     Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
     
    If Cells(i, 25).Value <> "todos iguales" And Cells(i, 25).Value <> "tiene 1 acumulados" And Cells(i, 25).Value <> "tiene 2 acumulados" And Cells(i, 25).Value <> "tiene 3 acumulados" And Cells(i, 25).Value <> "tiene 4 acumulados" Then
      If Cells(i, 25).Value = "tiene 5 acumulados" And Cells(i, 24).Value = "" Then
        mayor1 = Cells(i, 18).Value
        mayor2 = Cells(i, 20).Value
        
        'SE CALCULA LA DIFERENCIA ENTRE EL MAYOR1 Y MAYOR2
         
         dif = mayor1 - mayor2
         difporcentual = (dif * 100) / mayor1
         Cells(i, 22).Value = difporcentual
        
        'CALCULO EL POSIBLE SAC
         sac = mayor1 / 2
         Cells(i, 24).Value = sac
      End If
    End If
    
  Next i
   MsgBox "Proceso exitoso"
   Application.StatusBar = False
End Sub


Sub DIF_25_PORCIENTO()

Cells(1, 22).Value = "DIF DEL 25 %"

'SE CALCULA SI ENTRE EL 1ER MAYOR Y EL 2DO MAYOR EXISTE UNA DIFENCIA MAYOR DEL 25%

 mayor75 = 0
 mayor1 = 0
 
 limite = 5
 
  For i = 2 To limite
  
     Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
     
    If Cells(i, 23).Value <> "todos iguales" And Cells(i, 23).Value <> "tiene 1 acumulados" Then
    
        mayor1 = Cells(i, 18).Value
        
        'SE CALCULA EL 75% DEL 1ER MAYOR
        mayor75 = (mayor1 / 100) * 75
        
        'SE PREGUNTA SI EL SEGUNDO MAYOR ES MENOR QUE EL 75% DEL PRIMER MAYOR PARA SABER SI EXITE UNA DIFERENCIA MAYOR DEL 25% ENTRE EL 1ER MAYOR Y EL 2DO MAYOR
        If Cells(i, 20).Value < mayor75 Then
          Cells(i, 22).Value = "LA DIF ES MAYOR DEL 25%"
          Else
          Cells(i, 22).Value = "NO HAY DIF MAYOR DEL 25%"
        End If
        
    End If
    
  Next i
   MsgBox "Proceso exitoso"
   Application.StatusBar = False
End Sub
