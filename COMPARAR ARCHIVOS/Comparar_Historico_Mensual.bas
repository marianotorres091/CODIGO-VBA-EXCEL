Attribute VB_Name = "Módulo1"
    Sub Comprar_HISTORICO_MENSUAL()
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
    'Va la Hoja del Libro que se va a Abrir
    Set wsContenido = wbContenido.Worksheets("detalle")
    
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
    
    limite = nFilas
    For i = 2 To limite
    Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
       'libro que ya tengo abierto
                pos2 = i
                dni = Cells(i, 6).Value
                jur = Cells(i, 3).Value
                esc = Cells(i, 4).Value
                cuoc = Cells(i, 9).Value
                reaj = Cells(i, 10).Value
                unidad = Cells(i, 11).Value
                
                importe = Cells(i, 12).Value
                
                'Esta Función siempre que haya un 5, redondea hacia arriba osea si el num es 12,465 redondea a 12,47 y si es 12,455 redondea a 12,46
                importe = Application.WorksheetFunction.Round(importe, 2)
                
                vto = Cells(i, 13).Value
                band = False
             
              For j = 2 To nFilasCont
              
              'Va la hoja del libro que voy a abrir
               Set wsContenido = wbContenido.Worksheets("detalle")
               
                  importeRedond = wsContenido.Cells(j, 12).Value
                  
                 'VERIFICAR HACIA CUAL FUNCION USAR PARA REDONDEAR LOS DECIMALES
                  
                  'La funcion Round cuando hay un decimal 5, ejemplo 12,465 redondea hacia el par mas cercano osea 12,46 pero si el num es 12,455 redondea a 12,46
                  'importeRedond = Round(importeRedond, 2)
                  
                  'Esta Función siempre que haya un 5, redondea hacia arriba osea si el num es 12,465 redondea a 12,47 y si es 12,455 redondea a 12,46
                  'importeRedond = Application.WorksheetFunction.Round(importeRedond, 2)
                  
                 If wsContenido.Cells(j, 5).Value = dni Then
                    band = True
                    If wsContenido.Cells(j, 2).Value = jur Then
                         If wsContenido.Cells(j, 3).Value = esc Then
                            If wsContenido.Cells(j, 8).Value = cuoc Then
                               If wsContenido.Cells(j, 10).Value = reaj Then
                                  If wsContenido.Cells(j, 11).Value = unidad Then
                                    If importeRedond = importe Then
                                       If wsContenido.Cells(j, 15).Value = vto Then
                                         
                                               
                                                 pos = j
                                                 wsContenido.Cells(j, nColumnasCont + 1).Value = "MEN092020"
                                                
                                                 
                                                 
                                                     If wsContenido.Cells(j, nColumnasCont + 2).Value = "" Then
                                                           wsContenido.Cells(j, nColumnasCont + 2).Value = pos2
                                                        Else
                                                            pos4 = 3
                                                            Do While wsContenido.Cells(j, nColumnasCont + pos4).Value <> ""
                                                              pos4 = pos4 + 1
                                                            Loop
                                                            
                                                             If wsContenido.Cells(j, nColumnasCont + pos4).Value = "" Then
                                                               wsContenido.Cells(j, nColumnasCont + pos4).Value = pos2
                                                             End If
                                                      End If
                                                 
                                                  
                                                    'libro que ya tengo abierto
                                                  
                                                        Worksheets("Hoja1").Cells(i, nColumnas + 1).Value = "MEN092020"
                                                   
                                                    
                                                     If Worksheets("Hoja1").Cells(i, nColumnas + 2).Value = "" Then
                                                           Worksheets("Hoja1").Cells(i, nColumnas + 2).Value = pos
                                                        Else
                                                            pos3 = 3
                                                            Do While Worksheets("Hoja1").Cells(i, nColumnas + pos3).Value <> ""
                                                              pos3 = pos3 + 1
                                                            Loop
                                                            
                                                             If Worksheets("Hoja1").Cells(i, nColumnas + pos3).Value = "" Then
                                                               Worksheets("Hoja1").Cells(i, nColumnas + pos3).Value = pos
                                                             End If
                                                    End If
                                             
                                             
                                         
                                       End If
                                    End If
                                  End If
                               End If
                            End If
                         End If
                    End If
                    Else
                      If band Then
                        j = nFilasCont
                      End If
                 End If
                
                
            Next j
   
    Next i
     MsgBox "Proceso exitoso"
     Application.StatusBar = False
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub














    













    



