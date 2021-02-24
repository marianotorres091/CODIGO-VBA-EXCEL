Attribute VB_Name = "Módulo11"
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
    Set wbContenido = Application.Workbooks.Open("D:\MARIANO\CONCEPTO 246\246\246-04-2016-COBRADO--prueba.xlsx")
  
    
    'Activo el libro que estoy por abrir
    ThisWorkbook.Activate
    
    'va el nombre de la hoja del libro que voy a abrir
    Set wsContenido = wbContenido.Worksheets("A___HRG___Seleccion_de_Concepto")
    
    'va el nombre de la hoja del libro que ya tengo abierto
    Sheets("Año2016").Select
    
    'Calculo el número de filas de la hoja actual del libro que ya tengo abierto
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calculo el número de filas de la hoja de los cobrados
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    'en la primer columna vacia la nombro para ver cuales son los que que se encuentran en cobrados
    Cells(1, nColumnas + 1).Value = "IGUALES"
    
    
    For i = 2 To nFilas
        ' en las comillas va el mes que se va a pagar
        If Cells(i, 8).Value = "1" Then
            'Busco en el otro archivo el doc y marco en el historico-el valor de la celda que tomo es el del historico columna del dni
            valorDoc = Cells(i, 4).Value
            'defino el rango de busqueda del archivo de cobrado concatenando D con nFilasCont para que me quede "D-nº=D12" y el numero defino del final del rango
            rangoTemp = "F2:F" & nFilasCont
            Set resultado = wsContenido.Range(rangoTemp).Find(What:=valorDoc, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False, _
                        SearchFormat:=False)
            'Si el resultado de la búsqueda no es vacío
            If Not resultado Is Nothing Then
                
                
                'Estoy en la fila del historico correspondiente al dni encontrado en cobrados
                Cells(i, nColumnas + 1).Value = "COINCIDENCIA-NO PAGAR"
                
            Else
                Cells(i, nColumnas + 1).Value = "SI CORRESPONDE PAGAR"
            End If
        End If
    Next i
    
    MsgBox "Proceso exitosa"
    Exit Sub

End Sub




Sub Obtener_Carga()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim tempFecha As Date
    Dim i, RowNumber As Long
    Dim filaResultado As Long
    Dim filaCopia As Long
    Dim columnaCopia As Long
    Dim wbContenido As Worksheet
       

    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    

     Sheets.Add.Name = "RESULTADO"
 
    
    Cells(1, 1).Value = "PtaId"
    Cells(1, 2).Value = "JurId"
    Cells(1, 3).Value = "EscId"
    Cells(1, 4).Value = "Pref"
    Cells(1, 5).Value = "Doc"
    Cells(1, 6).Value = "Digito"
    Cells(1, 7).Value = "Nombres"
    Cells(1, 8).Value = "Couc"
    Cells(1, 9).Value = "Reajuste"
    Cells(1, 10).Value = "Unidades"
    Cells(1, 11).Value = "Importe"
    Cells(1, 12).Value = "Vto"
        
   filaResultado = 1
    
    
    For i = 2 To nFilas
       
                
                
                If Sheets("Hoja1").Cells(i, 27) = "" Then
                    filaResultado = filaResultado + 1
                    Sheets("RESULTADO").Cells(filaResultado, 1).Value = 0
                    Sheets("RESULTADO").Cells(filaResultado, 2).Value = Sheets("Hoja1").Cells(i, 1).Value
                    Sheets("RESULTADO").Cells(filaResultado, 3).Value = 2
                    Sheets("RESULTADO").Cells(filaResultado, 4).Value = 0
                    Sheets("RESULTADO").Cells(filaResultado, 5).Value = Sheets("Hoja1").Cells(i, 4).Value
                    Sheets("RESULTADO").Cells(filaResultado, 6).Value = 0
                    Sheets("RESULTADO").Cells(filaResultado, 7).Value = Sheets("Hoja1").Cells(i, 6).Value
                    Sheets("RESULTADO").Cells(filaResultado, 8).Value = Sheets("Hoja1").Cells(i, 22).Value
                    Sheets("RESULTADO").Cells(filaResultado, 9).Value = 1
                    Sheets("RESULTADO").Cells(filaResultado, 10).Value = 25
                    Sheets("RESULTADO").Cells(filaResultado, 11).Value = Sheets("Hoja1").Cells(i, 21).Value
                     Sheets("RESULTADO").Cells(filaResultado, 12).Value = "62017"
               End If
    Next i
    
                
       
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub

End Sub



Sub contar()
 RowNumber = ActiveSheet.Range("L132000").End(xlUp).Row
 MsgBox (RowNumber)
End Sub

Sub TOTAL()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i, TOTAL As Long
    

    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    For i = 2 To nFilas
      Sheets("RESULTADO").Cells(nFilas, 13).Value = Sheets("RESULTADO").Cells(nFilas, 13).Value + Sheets("RESULTADO").Cells(i, 11).Value
    Next i
   
    MsgBox "Proceso Exitoso el total es", , Cells(nFilas, 13).Value
End Sub


