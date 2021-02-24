Attribute VB_Name = "Módulo1"
Sub Actualizar_Drive()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim nColumnasCont As Long
    Dim i As Long
    Dim filaTotal As Long
    Dim nFilasError As Integer
    Dim columnaCopia As Long
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet
    Dim wsTotal As Excel.Worksheet


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
    Worksheets.Add
    ActiveSheet.Name = "ACT-VERIFICAR"
    Application.DisplayAlerts = True
    Set wsVerificar = Worksheets("ACT-VERIFICAR")
    
    'Va la Hoja del Libro que se va a Abrir
    Set wsContenido = wbContenido.Worksheets("Hoja2")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    MsgBox "VERIFICAR LA FECHA DE CUMPLIDA, LA COLUMNA DEL ARCHIVO DEL DRIVE DONDE VOY A CARGAR LA FECHA DE CUMPLIDA Y EL OPERADOR"
    
     'Encabezado Hoja Totales
    wsVerificar.Cells(1, 1).Value = "JURAS"
    wsVerificar.Cells(1, 2).Value = "AÑO"
    wsVerificar.Cells(1, 3).Value = "NUM"
    wsVerificar.Cells(1, 4).Value = "INGRESO"
    wsVerificar.Cells(1, 5).Value = "CUMPLIDA"
    wsVerificar.Cells(1, 6).Value = "OBSERVACIONES"
    filaTotal = 2
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    nColumnasCont = rangoCont.Columns.Count
    
    'Encabezados Hoja Informe
    Sheets("Hoja1").Cells(1, nColumnas + 1).Value = "ACT ECONTRADAS"
    Sheets("Hoja1").Cells(1, nColumnas + 2).Value = "OPERADORES ACTUALIZADOS"
    
    
    limite = nFilas
    For i = 2 To limite
    Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
       'libro que ya tengo abierto
                pos2 = i
                juras = Cells(i, 2).Value
                año = Cells(i, 3).Value
                num = Cells(i, 4).Value
                ultliq = Cells(i, 8).Value
                operador = Cells(i, 9).Value
              For j = 2 To nFilasCont
              
              'Va la hoja del libro que voy a abrir
               Set wsContenido = wbContenido.Worksheets("Hoja2")

                  
                 If wsContenido.Cells(j, 6).Value = juras Then
                    If wsContenido.Cells(j, 7).Value = año Then
                         If wsContenido.Cells(j, 8).Value = num Then
 
                    
                          If wsContenido.Cells(j, 2).Value = "" Then
                          
                           'Libro que ya tengo abierto
                            Worksheets("Hoja1").Cells(i, nColumnas + 1).Value = "Encontrada"
                            
                            If ultliq = "COM1020-11" Then
                               wsContenido.Cells(j, nColumnasCont + 1).Value = "15/11/2020"
                             Else
                              If ultliq = "MEN112020" Then
                               wsContenido.Cells(j, nColumnasCont + 1).Value = "31/11/2020"
                              End If
                            End If
                            
                          Else
                            'Libro que ya tengo abierto
                            Worksheets("Hoja1").Cells(i, nColumnas + 1).Value = "Encontrada-ya tiene FECH-CUMP"
                            
                            'Cargo en la hoja ACT-VERIFICAR
                            wsVerificar.Cells(filaTotal, 1).Value = wsContenido.Cells(j, 6).Value
                            wsVerificar.Cells(filaTotal, 2).Value = wsContenido.Cells(j, 7).Value
                            wsVerificar.Cells(filaTotal, 3).Value = wsContenido.Cells(j, 8).Value
                            wsVerificar.Cells(filaTotal, 4).Value = wsContenido.Cells(j, 1).Value
                            wsVerificar.Cells(filaTotal, 5).Value = wsContenido.Cells(j, 2).Value
                            wsVerificar.Cells(filaTotal, 6).Value = "VERIFICAR FECHA CUMPLIDA"
                            filaTotal = filaTotal + 1
                          End If
                            
                          'Verifico si los operadores son distintos y actualizo el drive
                          
                          If wsContenido.Cells(j, 16).Value <> operador Then
                          
                             If operador <> "" Then
                               wsContenido.Cells(j, nColumnasCont + 2).Value = operador
                             Else
                               wsContenido.Cells(j, nColumnasCont + 2).Value = "Sin operdador en Informe"
                             End If
                             
                            'Libro que ya tengo abierto
                              Worksheets("Hoja1").Cells(i, nColumnas + 2).Value = "Operador actualizado"
                          End If
                           
                           
                         End If
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



