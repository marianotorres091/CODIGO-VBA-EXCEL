Attribute VB_Name = "Módulo1"
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
    Set wbContenido = Application.Workbooks.Open("C:\Users\MARIANO\Desktop\Nueva carpeta (2)\PORTELA-PRUEBA.xlsx")
   
  
    
    'Activo el libro que estoy por abrir
    ThisWorkbook.Activate
    
    
    'va el nombre de la hoja del libro que voy a abrir
    Set wsContenido = wbContenido.Worksheets("Hoja1")
    
    'va el nombre de la hoja del libro que ya tengo abierto
    Sheets("HOJA1").Select
    Sheets("HOJA1").Cells(1, 16).Value = "ESTADO"
    
    'Calculo el número de filas de la hoja actual del libro que ya tengo abierto
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calculo el número de filas de la hoja de los cobrados
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    Workbooks("PORTELA-PRUEBA.xlsx").Activate
    Sheets("Hoja1").Cells(1, 18).Value = "ESTADO"
    Sheets("Hoja1").Cells(1, 19).Value = "Nº FILA ENCONTRADA "
    
    
    'en la primer columna vacia la nombro para ver cuales son los que que se encuentran en cobrados
    'Cells(1, nColumnas + 1).Value = "IGUALES"
    cont = 0
    cont2 = 0
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
        
        Workbooks("PORTELA-PRUEBA.xlsx").Activate
         If Sheets("Hoja1").Cells(j, 11).Value = cuoc Then
            If Sheets("Hoja1").Cells(j, 13).Value = rj Then
                 If Sheets("Hoja1").Cells(j, 14).Value = unidad Then
                    If Sheets("Hoja1").Cells(j, 15).Value = importe Then
                        If Sheets("Hoja1").Cells(j, 16).Value = vto Then
                          
                          Sheets("Hoja1").Cells(j, 18).Value = "ESTA"
                          cont = cont + 1
                          Workbooks("E6-2019-32576-A-PORTELA DELIA-JUDCIAL.CSV").Activate
                          Sheets("HOJA1").Cells(i, 17).Value = "ESTA"
                          Sheets("HOJA1").Cells(i, 18).Value = j
                           Workbooks("PORTELA-PRUEBA.xlsx").Activate
                          If Sheets("Hoja1").Cells(j, 19).Value = "" Then
                              Sheets("Hoja1").Cells(j, 19).Value = V1
                              Sheets("Hoja1").Cells(j, 20).Value = corresponde
                              Else
                              Sheets("Hoja1").Cells(j, 21).Value = V1
                              Sheets("Hoja1").Cells(j, 22).Value = corresponde
                              cont2 = cont2 + 1
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
     MsgBox cont
     MsgBox cont2
End Sub

Sub Comparar_Archivos3()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i, cuoc, rj, unidad, importe, vto As Long
   

    
    
    For i = 2 To 592
       
        Cells(i, 16).Value = "buscado"
        corresponde = Cells(i, 14).Value
        V1 = i
      
      For j = 2 To 649
        
       
         If Cells(j, 25).Value = Cells(i, 8).Value Then
            If Cells(j, 26).Value = Cells(i, 9).Value Then
                 If Cells(j, 27).Value = Cells(i, 10).Value Then
                    If Cells(j, 28).Value = Cells(i, 11).Value Then
                        If Cells(j, 29).Value = Cells(i, 12).Value Then
                          
                          Cells(j, 31).Value = "ESTA"
                         
                          If Cells(j, 33).Value = "" Then
                              Cells(j, 33).Value = V1
                              Cells(j, 33).Value = corresponde
                              Else
                              Cells(j, 34).Value = V1
                              Cells(j, 35).Value = corresponde
                              
                              Cells(i, 17).Value = "falta"
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
Sub Comparar_Archivos()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i As Long
    Dim filaCopia As Long
    Dim nFilasError As Integer
    Dim columnaCopia As Long
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet


    
    'abro el libro con el que voy a comparar con el Historico
    Set wbContenido = Application.Workbooks.Open("C:\Users\MARIANO\Desktop\Nueva carpeta\E6-2019-32576-A-PORTELA DELIA-JUDCIAL-CON DUPLICADOS.xlsx")
  
    
    'Activo el libro que estoy por abrir
    ThisWorkbook.Activate
    
    
    'va el nombre de la hoja del libro que voy a abrir
    Set wsContenido = wbContenido.Worksheets("HOJA1")
    
    'va el nombre de la hoja del libro que ya tengo abierto
    Sheets("HOJA1").Select
    Sheets("HOJA1").Cells(1, 16).Value = "ESTADO"
    
    'Calculo el número de filas de la hoja actual del libro que ya tengo abierto
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calculo el número de filas de la hoja de los cobrados
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    Workbooks("E6-2019-32576-A-PORTELA DELIA-JUDCIAL-CON DUPLICADOS.xlsx").Activate
    Sheets("HOJA1").Cells(1, 16).Value = "ESTADO"
    Sheets("HOJA1").Cells(1, 17).Value = "Nº FILA ENCONTRADA "
    
    
    'en la primer columna vacia la nombro para ver cuales son los que que se encuentran en cobrados
    'Cells(1, nColumnas + 1).Value = "IGUALES"
    
    
    For i = 2 To nFilas
        Workbooks("E6-2019-32576-A-PORTELA DELIA-JUDCIAL.CSV").Activate
        Cells(i, 16).Value = "buscado"
        V1 = i
        V2 = Cells(i, 8).Value
        V3 = Cells(i, 9).Value
        V4 = Cells(i, 10).Value
        V5 = Cells(i, 11).Value
        V6 = Cells(i, 12).Value
      For j = 2 To nFilas
        
        Workbooks("E6-2019-32576-A-PORTELA DELIA-JUDCIAL-CON DUPLICADOS.xlsx").Activate
         If Sheets("HOJA1").Cells(j, 8).Value = V2 Then
            If Sheets("HOJA1").Cells(j, 9).Value = V3 Then
                 If Sheets("HOJA1").Cells(j, 10).Value = V4 Then
                    If Sheets("HOJA1").Cells(j, 11).Value = V5 Then
                        If Sheets("HOJA1").Cells(j, 12).Value = V6 Then
                          
                          Sheets("HOJA1").Cells(j, 16).Value = "ESTA"
                         
                          If Sheets("HOJA1").Cells(j, 17).Value = "" Then
                              Sheets("HOJA1").Cells(j, 17).Value = V1
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
Sub BuscarDuplicados()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i As Long
    

    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count

    For i = 2 To nFilas
       Cells(i, 20).Value = "buscado"
       Cells(i, 21).Value = i
      For j = 2 To nFilas
        If Cells(j, 20).Value <> "buscado" Then
            If Cells(i, 11).Value = Cells(j, 11).Value Then
              If Cells(i, 8).Value = Cells(j, 8).Value Then
               If Cells(i, 9).Value = Cells(j, 9).Value Then
                If Cells(i, 10).Value = Cells(j, 10).Value Then
                 If Cells(i, 12).Value = Cells(j, 12).Value Then
                    Cells(i, 22).Value = "repetido"
                    Cells(j, 23).Value = i
                 End If
                End If
               End If
              End If
            End If
        End If
      Next j
      Cells(i, 20).Value = " "
    Next i
   
    MsgBox "Proceso Exitoso"
End Sub





