Attribute VB_Name = "Módulo1"
Sub Comparar_dif_Archivos()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i As Long
    Dim filaCopia As Long
    Dim nFilasError, inicio, limite As Integer
    Dim columnaCopia As Long
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet


    
    'abro el libro con el que voy a comparar con el Historico
    Set wbContenido = Application.Workbooks.Open("D:\TRABAJO\CARGAR 2020\MAYO 2020\COTEJAR\ULTIMO - HISTORICO 1-actualizado 20-05-2020 -prueba-.xlsx")
  
    
    'Activo el libro que estoy por abrir
    ThisWorkbook.Activate
    
    
    'va el nombre de la hoja del libro que voy a abrir
    Set wsContenido = wbContenido.Worksheets("Hoja1")
  
    
    'va el nombre de la hoja del libro que ya tengo abierto
    Sheets("HISTORICO").Select
   
    
    'Calculo el número de filas de la hoja actual del libro que ya tengo abierto
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    'Sheets("HISTORICO").Cells(1, 35).Value = "ESTA"
    'Sheets("HISTORICO").Cells(1, 36).Value = "POSICIÓN-HISTORICO"
    
    
    'Calculo el número de filas de la hoja de los cobrados
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    'el libro que voy a abrir
    Workbooks("ULTIMO - HISTORICO 1-actualizado 20-05-2020 -prueba-.xlsx").Activate
     'Sheets("Hoja1").Cells(1, 30).Value = "OBSERVACION"
    
    
    'en la primer columna vacia la nombro para ver cuales son los que que se encuentran en cobrados
    'Cells(1, nColumnas + 1).Value = "IGUALES"
    
    
    limite = 3672
    For i = 2 To limite
    Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
       'libro que ya tengo habierto el historico
        Workbooks("CUOTAS 52020 - FINAL 1.xlsx").Activate
        
        pos2 = i
        dni = Cells(i, 5).Value
        jur = Cells(i, 2).Value
        esc = Cells(i, 3).Value
        cuoc = Cells(i, 8).Value
        reaj = Cells(i, 9).Value
        unidad = Cells(i, 10).Value
        importe = Cells(i, 11).Value
        vto = Cells(i, 12).Value
        act = Cells(i, 14).Value
        cuota = Cells(i, 13).Value
        liquidado = Cells(i, 24).Value
        pago = Cells(i, 27).Value
        totalcuota = Cells(i, 28).Value
        habilito = Cells(i, 29).Value
        partir = Cells(i, 30).Value
        coupend = Cells(i, 31).Value
        esta = Cells(i, 33).Value
        nom = Cells(i, 7).Value
      For j = 2 To 6622
         'el libro que voy a abrir
        Workbooks("ULTIMO - HISTORICO 1-actualizado 20-05-2020 -prueba-.xlsx").Activate
         
         If Sheets("Hoja1").Cells(j, 5).Value = dni Then
            If Sheets("Hoja1").Cells(j, 2).Value = jur Then
                 If Sheets("Hoja1").Cells(j, 3).Value = esc Then
                    If Sheets("Hoja1").Cells(j, 8).Value = cuoc Then
                        If Sheets("Hoja1").Cells(j, 9).Value = reaj Then
                          If Sheets("Hoja1").Cells(j, 10).Value = unidad Then
                            If Sheets("Hoja1").Cells(j, 11).Value = importe Then
                               If Sheets("Hoja1").Cells(j, 12).Value = vto Then
                                  If Sheets("Hoja1").Cells(j, 14).Value = act Then
                                    If Sheets("Hoja1").Cells(j, 13).Value = cuota Then
                                        If Sheets("Hoja1").Cells(j, 21).Value = liquidado Then
                                           If Sheets("Hoja1").Cells(j, 24).Value = pago Then
                                               If Sheets("Hoja1").Cells(j, 25).Value = totalcuota Then
                                                  If Sheets("Hoja1").Cells(j, 26).Value = habilito Then
                                                      If Sheets("Hoja1").Cells(j, 27).Value = partir Then
                                                         If Sheets("Hoja1").Cells(j, 28).Value = coupend Then
                                                                 'pos = j
                                                                 
                                                                 Sheets("Hoja1").Cells(j, 37).Value = jur
                                                                 Sheets("Hoja1").Cells(j, 38).Value = esc
                                                                 Sheets("Hoja1").Cells(j, 40).Value = dni
                                                                 Sheets("Hoja1").Cells(j, 42).Value = nom
                                                                 Sheets("Hoja1").Cells(j, 43).Value = cuoc
                                                                 Sheets("Hoja1").Cells(j, 44).Value = reaj
                                                                 Sheets("Hoja1").Cells(j, 45).Value = unidad
                                                                 Sheets("Hoja1").Cells(j, 46).Value = importe
                                                                 Sheets("Hoja1").Cells(j, 47).Value = vto
                                                                 Sheets("Hoja1").Cells(j, 48).Value = cuota
                                                                 Sheets("Hoja1").Cells(j, 49).Value = act
                                                                 Sheets("Hoja1").Cells(j, 56).Value = liquidado
                                                                 Sheets("Hoja1").Cells(j, 59).Value = pago
                                                                 Sheets("Hoja1").Cells(j, 60).Value = totalcuota
                                                                 Sheets("Hoja1").Cells(j, 61).Value = habilito
                                                                 Sheets("Hoja1").Cells(j, 62).Value = partir
                                                                 Sheets("Hoja1").Cells(j, 63).Value = coupend
                                                                 'libro que ya tengo habierto el historico
                                                                 Workbooks("CUOTAS 52020 - FINAL 1.xlsx").Activate
                                                                 Sheets("HISTORICO").Cells(i, 42).Value = "copiado"
                                                                 'Sheets("HISTORICO").Cells(i, 24).Value = "COM 04-20-08"
                                                                 
                                                                  'If Sheets("HISTORICO").Cells(i, 41).Value = "" Then
                                                                      'Sheets("HISTORICO").Cells(i, 41).Value = "iguales"
                                                                     
                                                                   'End If
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









