Attribute VB_Name = "Módulo1"
Sub REPORTE()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i, cuoc, rj, unidad, importe, vto, cont, cont2, fecha As Long
    Dim filaCopia As Long
    Dim nFilasError As Integer
    Dim columnaCopia As Long
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet

       'abro el libro con el que voy a comparar con el Historico
    Set wbContenido = Application.Workbooks.Open("C:\Users\MARIANO\Desktop\MESA ENTRADA\ENVIADOS.xlsx")
   
  
    
    'Activo el libro que estoy por abrir
    ThisWorkbook.Activate
    
    
    'va el nombre de la hoja del libro que voy a abrir
    Set wsContenido = wbContenido.Worksheets("HOJA1")
    
    'va el nombre de la hoja del libro que ya tengo abierto
    Sheets("Hoja1").Select
    Sheets("Hoja1").Cells(1, 1).Value = "FECHA-ENVÍO"
    
    'Calculo el número de filas de la hoja actual del libro que ya tengo abierto
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calculo el número de filas de la hoja de los cobrados
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    
    
    'en la primer columna vacia la nombro para ver cuales son los que que se encuentran en cobrados
    'Cells(1, nColumnas + 1).Value = "IGUALES"
    
    For i = 2 To 5
        Workbooks("Actuaciones.xlsx").Activate
        Cells(i, 25).Value = "buscado"
        V1 = i
        act = Cells(i, 2).Value
      For j = 2 To 9324
        
        Workbooks("ENVIADOS.xlsx").Activate
        
                        If Sheets("HOJA1").Cells(j, 1).Value = act Then
                          fecha = Sheets("HOJA1").Cells(j, 2).Value
                          Sheets("HOJA1").Cells(j, 5).Value = "ESTA"
                          
                          Workbooks("Actuaciones.xlsx").Activate
                          Sheets("Hoja1").Cells(i, 1).Value = fecha
                          Sheets("Hoja1").Cells(i, 22).Value = "ESTA"
                          Sheets("Hoja1").Cells(i, 23).Value = j
                           Workbooks("ENVIADOS.xlsx").Activate
                        End If
             
          
          
      Next j
       
    Next i
     MsgBox "Proceso exitosa"
   
End Sub

