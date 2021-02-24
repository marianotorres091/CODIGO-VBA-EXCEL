Attribute VB_Name = "Módulo1"
    Sub Comprar_Archivos_Diferentes_con_montos_totales()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim nColumnasCont As Long
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
    'Worksheets.Add
    'ActiveSheet.Name = "Errores"
    Application.DisplayAlerts = True
    'Set wsError = Worksheets("Errores")
    Set wsContenido = wbContenido.Worksheets("CUOTAS BAJADAS EN MAYO - 13-07 ")
    
    'Regresa el control a la hoja de origen
    Sheets("faltan archivos").Select
    
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    nColumnasCont = rangoCont.Columns.Count
    
    limite = nFilas
    For i = 2 To limite
    Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
       'libro que ya tengo abierto
  
    
                pos2 = i
                dni = Cells(i, 2).Value
                'jur = Cells(i, 2).Value
                'esc = Cells(i, 3).Value
                cuoc = Cells(i, 10).Value
                'reaj = Cells(i, 10).Value
                unidad = Cells(i, 18).Value
                'importe = Cells(i, 12).Value
                'vto = Cells(i, 14).Value
            
              For j = 2 To nFilasCont
                 'el libro que voy a abrir
               Set wsContenido = wbContenido.Worksheets("CUOTAS BAJADAS EN MAYO - 13-07 ")
               
                  'importeRedond = wsContenido.Cells(j, 12).Value
                  'importeRedond = Round(importeRedond, 2)
              
                 If wsContenido.Cells(j, 5).Value = dni Then
                    'If wsContenido.Cells(j, 3).Value = jur Then
                         'If wsContenido.Cells(j, 3).Value = esc Then
                            If wsContenido.Cells(j, 8).Value = cuoc Then
                               'If wsContenido.Cells(j, 10).Value = reaj Then
                                  If wsContenido.Cells(j, 10).Value = unidad Then
                                    'If importeRedond = importe Then
                                       'If wsContenido.Cells(j, 13).Value = vto Then
                                         
                                               
                                                 pos = j
                                                 
                                                If wsContenido.Cells(j, nColumnasCont + 1).Value = "" Then
                                                 wsContenido.Cells(j, nColumnasCont + 1).Value = "existe en arch faltantes "
                                                End If
                                                 
                                                 
                                                     If wsContenido.Cells(j, nColumnasCont + 2).Value = "" Then
                                                           wsContenido.Cells(j, nColumnasCont + 2).Value = pos2
                                                        Else
                                                            pos4 = 3
                                                            Do While wsContenido.Cells(j, nColumnasCont + pos4).Value <> ""
                                                              pos4 = pos4 + 1
                                                            Loop
                                                            
                                                             If wsContenido.Cells(j, nColumnasCont + pos4).Valu = "" Then
                                                               wsContenido.Cells(j, nColumnasCont + pos4).Valu = pos2
                                                             End If
                                                      End If
                                                  
                                                  
                                                 
                                                  
                                                    'libro que ya tengo abierto
                                                  
                                                        Worksheets("faltan archivos").Cells(i, nColumnas + 1).Value = "esta en cuotas bajadas"
                                                   
                                                    
                                                        Worksheets("faltan archivos").Cells(i, nColumnas + 2).Value = Worksheets("faltan archivos").Cells(i, nColumnas + 2).Value + wsContenido.Cells(j, 11).Value
                                                     
                                             
                                         
                                       'End If
                                    'End If
                                  End If
                               'End If
                               Else
                               If wsContenido.Cells(j, 8).Value = 316 Then
                                 Worksheets("faltan archivos").Cells(i, nColumnas + 2).Value = Worksheets("faltan archivos").Cells(i, nColumnas + 2).Value + wsContenido.Cells(j, 11).Value
                               End If
                            End If
                          'End If
                    'End If

                 End If
                
                
            Next j
   
    Next i
     MsgBox "Proceso exitoso"
     Application.StatusBar = False
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub










