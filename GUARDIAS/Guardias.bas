Attribute VB_Name = "Módulo1"
Sub DNI_iguales_TipoProf_diferentes_paso1()
    Dim nFilas As Long
    Dim nColumnas As Long
    'Dim i As Integer
    Dim rango As Range
    Dim band As Double
    
    
    'ATECION!!! PRIMERO APLICAR EL ORDENAMIENTO POR DNI Y UNA VEZ QUE APLIQUE VUELVO A APLICAR EL ORDENAMIENTO POR TIPOPROF!!!
    'SI APLICO EL ORDENAMIENTO POR DNI Y TIPOPROF JUNTOS NO FUNCIONA EL CODIGO
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    Sheets.Add.Name = "ERRORES"
    
    filaResultado = 1
    
    Cells(filaResultado, 1).Value = "CUOF"
    Cells(filaResultado, 2).Value = "ANEXO"
    Cells(filaResultado, 3).Value = "AÑO"
    Cells(filaResultado, 4).Value = "MES"
    Cells(filaResultado, 5).Value = "DNI"
    Cells(filaResultado, 6).Value = "APELLIDO Y NOMBRE"
    
    Sheets("Hoja1").Select
    
    band = False
    limite = (nFilas - 1)
    For i = 2 To limite
     Application.StatusBar = Format(i / limite, "0.0%") & " " & "Completo"
     
    
     If Cells(i, nColumnas + 1).Value <> "DNI= - TIPOPROF DIST" Then
        
        cuof = Cells(i, 1).Value
        Doc = Cells(i, 5).Value
        año = Cells(i, 3).Value
        mes = Cells(i, 4).Value
        anexo = Cells(i, 2).Value
        nombre = Cells(i, 6).Value
        tipoProf = Cells(i, 7).Value
        
        For j = (i + 1) To nFilas
            If Doc = Cells(j, 5).Value And tipoProf <> Cells(j, 7).Value Then
             If tipoProf = "A" Then
               If Cells(j, 7).Value <> "D" Then
                    Sheets("ERRORES").Select
                     'Calcular el número de filas de la hoja ERRORES
                        Set rango = ActiveSheet.UsedRange
                        nFilasE = rango.Rows.Count
                        nColumnasE = rango.Columns.Count
                        
                     For e = 1 To nFilasE
                      If Cells(e, 5).Value = Doc Then
                       band = True
                      End If
                     Next e
                     
                     If band = False Then
                        filaResultado = filaResultado + 1
                        Cells(filaResultado, 1).Value = cuof
                        Cells(filaResultado, 2).Value = anexo
                        Cells(filaResultado, 3).Value = año
                        Cells(filaResultado, 4).Value = mes
                        Cells(filaResultado, 5).Value = Doc
                        Cells(filaResultado, 6).Value = nombre
                      Else
                      band = False
                    End If
                    
                    Sheets("Hoja1").Select
                    
                    Cells(j, 5).Interior.Color = RGB(240, 243, 121)
                    Cells(i, 5).Interior.Color = RGB(240, 243, 121)
                    Cells(j, nColumnas + 1).Value = "DNI= - TIPOPROF DIST"
                    If Cells(j, nColumnas + 2).Value = "" Then
                        Cells(j, nColumnas + 2).Value = i
                        Cells(i, nColumnas + 2).Value = i
                       Else
                        Cells(i, nColumnas + 2).Value = Cells(j, nColumnas + 2).Value
                     End If
                    Cells(i, nColumnas + 1).Value = "DNI= - TIPOPROF DIST"
               End If
               Else
                    Sheets("ERRORES").Select
                     'Calcular el número de filas de la hoja ERRORES
                        Set rango = ActiveSheet.UsedRange
                        nFilasE = rango.Rows.Count
                        nColumnasE = rango.Columns.Count
                        
                     For e = 1 To nFilasE
                      If Cells(e, 5).Value = Doc Then
                       band = True
                      End If
                     Next e
                     
                     If band = False Then
                        filaResultado = filaResultado + 1
                        Cells(filaResultado, 1).Value = cuof
                        Cells(filaResultado, 2).Value = anexo
                        Cells(filaResultado, 3).Value = año
                        Cells(filaResultado, 4).Value = mes
                        Cells(filaResultado, 5).Value = Doc
                        Cells(filaResultado, 6).Value = nombre
                      Else
                      band = False
                    End If
                    
                    Sheets("Hoja1").Select
                    
                    Cells(j, 5).Interior.Color = RGB(240, 243, 121)
                    Cells(i, 5).Interior.Color = RGB(240, 243, 121)
                    Cells(j, nColumnas + 1).Value = "DNI= - TIPOPROF DIST"
                     If Cells(j, nColumnas + 2).Value = "" Then
                        Cells(j, nColumnas + 2).Value = i
                        Cells(i, nColumnas + 2).Value = i
                       Else
                        Cells(i, nColumnas + 2).Value = Cells(j, nColumnas + 2).Value
                     End If
                    Cells(i, nColumnas + 1).Value = "DNI= - TIPOPROF DIST"
             End If
            End If
        Next j
      End If
  
    Next i
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Application.StatusBar = False
End Sub
    Sub Verificar_tipoprof_en_HistoricoGuardias_paso2()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim nColumnasCont As Long
    Dim band As Double
    Dim i As Long
    Dim filaCopia As Long
    Dim nFilasError As Integer
    Dim columnaCopia As Long
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet


    'Indicar el libro de excel CONTENIDO y control de errores
    contenido = InputBox("Ingrese el nombre del archivo:", "Abrir", "Archivo.xlsx")
    If contenido <> "" Then
       ' On Error GoTo ControlErrorOpen
        Set wbContenido = Workbooks.Open(ActiveWorkbook.Path & "\" & contenido)
    Else
        Exit Sub
    End If
    
    'Activar este libro
    ThisWorkbook.Activate
    
    Application.DisplayAlerts = False
    Application.DisplayAlerts = True
    Set wsContenido = wbContenido.Worksheets("HISTORICO")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    nColumnasCont = rangoCont.Columns.Count
    
    nColumnas = nColumnas + 1
    nColumnasCont = nColumnasCont + 1
    limite = nFilas
    band = False
    
    For i = 2 To limite
     Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
       'libro que ya tengo abierto
     If Cells(i, 12).Value <> "DNI= - TIPOPROF DIST" Then
               dni = Cells(i, 5).Value
               tipoProf = Cells(i, 7).Value
              
              
            For j = 2 To nFilasCont
              
                 'el libro que voy a abrir
               Set wsContenido = wbContenido.Worksheets("HISTORICO")
                
                 If wsContenido.Cells(j, 1).Value = dni Then
                  band = True
                   If wsContenido.Cells(j, 3).Value = "A" Then
                      If tipoProf = "D" Then
                            'Regresa el control a la hoja de origen
                             Sheets("Hoja1").Select
        
                            'libro que ya tengo abierto
                             Worksheets("Hoja1").Cells(i, nColumnas).Value = "IGUALES BD=A y RegNuev=D"
                            
                        Else
                          If tipoProf = "A" Then
                           'Regresa el control a la hoja de origen
                             Sheets("Hoja1").Select
        
                            'libro que ya tengo abierto
                             Worksheets("Hoja1").Cells(i, nColumnas).Value = "IGUALES"
                            
                           Else
                             Worksheets("Hoja1").Cells(i, nColumnas).Value = "TipoProfDistinto"
                              Set wsContenido = wbContenido.Worksheets("HISTORICO")
                                  wsContenido.Cells(j, nColumnasCont).Value = "VERIFICAR TIPOPROF"
                               
                           End If
                       End If
                     Else
                       If wsContenido.Cells(j, 3).Value = "D" Then
                          If tipoProf = "A" Then
                             'Regresa el control a la hoja de origen
                             Sheets("Hoja1").Select
        
                            'libro que ya tengo abierto
                             Worksheets("Hoja1").Cells(i, nColumnas).Value = "IGUALES BD=D y RegNuev=A"
                             
                           Else
                              If tipoProf = "D" Then
                               'Regresa el control a la hoja de origen
                                 Sheets("Hoja1").Select
            
                                'libro que ya tengo abierto
                                 Worksheets("Hoja1").Cells(i, nColumnas).Value = "IGUALES BD=A y RegNuev=D"
                                
                               Else
                                 Worksheets("Hoja1").Cells(i, nColumnas).Value = "TipoProfDistinto"
                                  Set wsContenido = wbContenido.Worksheets("HISTORICO")
                                  wsContenido.Cells(j, nColumnasCont).Value = "VERIFICAR TIPOPROF"
                              
                              End If
                           End If
                    
                        Else
                        
                            If wsContenido.Cells(j, 3).Value = tipoProf Then
                                'Regresa el control a la hoja de origen
                                 Sheets("Hoja1").Select
            
                                'libro que ya tengo abierto
                                 Worksheets("Hoja1").Cells(i, nColumnas).Value = "IGUALES"
                                
                             Else
                                 Worksheets("Hoja1").Cells(i, nColumnas).Value = "TipoProfDistinto"
                                 Set wsContenido = wbContenido.Worksheets("HISTORICO")
                                  wsContenido.Cells(j, nColumnasCont).Value = "VERIFICAR TIPOPROF"
                              
                            End If
                       End If
                   End If
                 End If
            Next j
           
            If band = False Then
            'Regresa el control a la hoja de origen
             Sheets("Hoja1").Select
            
             'libro que ya tengo abierto
              Worksheets("Hoja1").Cells(i, nColumnas).Value = "NUEVO A AGREGAR"
             Else
             band = False
            End If
     End If
    Next i
     MsgBox "Proceso exitoso"
     Application.StatusBar = False
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub
    Sub Comparar_Guardias_HistoricoGuardias_cargarNuevos_paso3()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i As Long
    Dim filaCopia As Long
    Dim nFilasError As Integer
    Dim band As Double
    Dim columnaCopia As Long
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet


    'Indicar el libro de excel CONTENIDO y control de errores
    contenido = InputBox("Ingrese el nombre del archivo:", "Abrir", "Archivo.xlsx")
    If contenido <> "" Then
       ' On Error GoTo ControlErrorOpen
        Set wbContenido = Workbooks.Open(ActiveWorkbook.Path & "\" & contenido)
    Else
        Exit Sub
    End If
    
    'Activar este libro
    ThisWorkbook.Activate
    
    Application.DisplayAlerts = False
    Application.DisplayAlerts = True
    Set wsContenido = wbContenido.Worksheets("HISTORICO")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
   
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    band = False
    nColumnas = nColumnas + 1
    limite = nFilas
    For i = 2 To limite
     Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
       'libro que ya tengo abierto
     If Sheets("Hoja1").Cells(i, 14).Value = "NUEVO A AGREGAR" Then
    
            
               cuof = Cells(i, 1).Value
               anexo = Cells(i, 2).Value
               año = Cells(i, 3).Value
               mes = Cells(i, 4).Value
               dni = Cells(i, 5).Value
               nombre = Cells(i, 6).Value
               tipoProf = Cells(i, 7).Value
               cuoc = Cells(i, 8).Value
               horas = Cells(i, 9).Value
               servicio = Cells(i, 10).Value
              
            For j = 2 To nFilasCont
              
                 'el libro que voy a abrir
               Set wsContenido = wbContenido.Worksheets("HISTORICO")
                 
                 If wsContenido.Cells(j, 1).Value = dni Then
                    band = True
                    If cuoc = 275 Then
                      wsContenido.Cells(j, 8).Value = wsContenido.Cells(j, 8).Value + horas
                      Else
                       If cuoc = 276 Then
                         wsContenido.Cells(j, 9).Value = wsContenido.Cells(j, 9).Value + horas
                         Else
                          wsContenido.Cells(j, 10).Value = wsContenido.Cells(j, 10).Value + horas
                       End If
                    End If
                    
                   If wsContenido.Cells(j, 2).Value = "" Then
                    wsContenido.Cells(j, 2).Value = nombre
                    wsContenido.Cells(j, 3).Value = tipoProf
                    wsContenido.Cells(j, 4).Value = cuof
                    wsContenido.Cells(j, 5).Value = anexo
                    wsContenido.Cells(j, 6).Value = año
                    wsContenido.Cells(j, 7).Value = mes
                    wsContenido.Cells(j, 11).Value = servicio
                   End If
                    'libro que ya tengo abierto
                   
                     Worksheets("Hoja1").Cells(i, nColumnas).Value = "ENCONTRADO"
                  Else
                  band = False
                 End If
                
            Next j
            
            If band = False Then
            
              nFilasCont = nFilasCont + 1
             
              If wsContenido.Cells(nFilasCont, 1).Value = "" Then
              
                wsContenido.Cells(nFilasCont, 8).Value = 0
                wsContenido.Cells(nFilasCont, 9).Value = 0
                wsContenido.Cells(nFilasCont, 10).Value = 0
                wsContenido.Cells(nFilasCont, 1).Value = dni
                      
                      If cuoc = 275 Then
                        wsContenido.Cells(nFilasCont, 8).Value = wsContenido.Cells(nFilasCont, 8).Value + horas
                        Else
                         If cuoc = 276 Then
                           wsContenido.Cells(nFilasCont, 9).Value = wsContenido.Cells(nFilasCont, 9).Value + horas
                           Else
                            wsContenido.Cells(nFilasCont, 10).Value = wsContenido.Cells(nFilasCont, 10).Value + horas
                         End If
                      End If
                      
                     
                      wsContenido.Cells(nFilasCont, 2).Value = nombre
                      wsContenido.Cells(nFilasCont, 3).Value = tipoProf
                      wsContenido.Cells(nFilasCont, 4).Value = cuof
                      wsContenido.Cells(nFilasCont, 5).Value = anexo
                      wsContenido.Cells(nFilasCont, 6).Value = año
                      wsContenido.Cells(nFilasCont, 7).Value = mes
                      wsContenido.Cells(nFilasCont, 11).Value = servicio
                      
                     'libro que ya tengo abierto
                     Sheets("Hoja1").Select
                     Sheets("Hoja1").Cells(i, nColumnas).Value = "NUEVO AGREGADO"
                     
              End If
            End If
     End If
    Next i
     MsgBox "Proceso exitoso"
     Application.StatusBar = False
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub

   

