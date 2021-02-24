Attribute VB_Name = "Module1"
Sub PORTELA()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i, cuoc, rj, unidad, importe, vto, cont, cont2 As Long
    Dim filaCopia As Long
    Dim nFilasError As Integer
    Dim columnaCopia As Long
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet

       'abro el libro con el que voy a comparar con el Historico
    Set wbContenido = Application.Workbooks.Open("C:\Users\MARIANO\Desktop\PORTELA DELIA\CALCULADOS\ENVIADO A RAUL\MANDE A RAUL ULTIMO\PORTELA.xlsx")
   
  
    
    'Activo el libro que estoy por abrir
    ThisWorkbook.Activate
    
    
    'va el nombre de la hoja del libro que voy a abrir
    Set wsContenido = wbContenido.Worksheets("Hoja1")
    
    'va el nombre de la hoja del libro que ya tengo abierto
    Sheets("PLANILLA PORTELA DELIA INTERES ").Select
    Sheets("PLANILLA PORTELA DELIA INTERES ").Cells(1, 16).Value = "ESTADO"
    
    'Calculo el número de filas de la hoja actual del libro que ya tengo abierto
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calculo el número de filas de la hoja de los cobrados
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    Workbooks("PORTELA.xlsx").Activate
    Sheets("Hoja1").Cells(1, 18).Value = "ESTADO"
    Sheets("Hoja1").Cells(1, 19).Value = "Nº FILA ENCONTRADA "
    
    
    'en la primer columna vacia la nombro para ver cuales son los que que se encuentran en cobrados
    'Cells(1, nColumnas + 1).Value = "IGUALES"
    
    For i = 2 To 592
        Workbooks("E6-2019-32576-A-PORTELA DELIA-JUDCIAL.CSV").Activate
        Cells(i, 16).Value = "buscado"
        corresponde = Cells(i, 14).Value
        V1 = i
        cuoc = Cells(i, 8).Value
        rj = Cells(i, 9).Value
        unidad = Cells(i, 10).Value
        importe = Cells(i, 11).Value
        vto = Cells(i, 12).Value
      For j = 2 To 649
        
        Workbooks("PORTELA.xlsx").Activate
         If Sheets("Hoja1").Cells(j, 11).Value = cuoc Then
            If Sheets("Hoja1").Cells(j, 13).Value = rj Then
                 If Sheets("Hoja1").Cells(j, 14).Value = unidad Then
                    If Sheets("Hoja1").Cells(j, 15).Value = importe Then
                        If Sheets("Hoja1").Cells(j, 16).Value = vto Then
                          
                          Sheets("PORTELA").Cells(j, 18).Value = "ESTA"
                         
                          Workbooks("E6-2019-32576-A-PORTELA DELIA-JUDCIAL.CSV").Activate
                          Sheets("HOJA1").Cells(i, 17).Value = "ESTA"
                          Sheets("HOJA1").Cells(i, 18).Value = j
                           Workbooks("PORTELA.xlsx").Activate
                          If Sheets("Hoja1").Cells(j, 19).Value = "" Then
                              Sheets("Hoja1").Cells(j, 19).Value = V1
                              Sheets("Hoja1").Cells(j, 20).Value = corresponde
                              Else
                              Sheets("Hoja1").Cells(j, 21).Value = V1
                              Sheets("Hoja1").Cells(j, 22).Value = corresponde
                              
                              Workbooks("E6-2019-32576-A-PORTELA DELIA-JUDCIAL.CSV").Activate
                              Cells(i, 19).Value = "falta"
                          End If
                      
                        End If
                     End If
                  End If
            End If
         End If
          
          
      Next j
       
    Next i
     MsgBox "Proceso exitosa"
     
End Sub

