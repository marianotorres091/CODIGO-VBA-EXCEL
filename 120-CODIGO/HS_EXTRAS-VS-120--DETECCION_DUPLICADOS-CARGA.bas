Attribute VB_Name = "Módulo1"
Sub Comparar_Archivos120_324()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i As Long
    Dim filaCopia As Long
    Dim nFilasError As Integer
    Dim columnaCopia As Long
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet


    
    'abro el libro de hs extras que voy a comparar con el 120 del mensual
     'Indicar el libro de excel CONTENIDO y control de errores
    contenido = InputBox("Ingrese el nombre del archivo:", "Abrir", "Archivo.xlsx")
    If contenido <> "" Then
       ' On Error GoTo ControlErrorOpen
        Set wbContenido = Workbooks.Open(ActiveWorkbook.Path & "\" & contenido)
    Else
        Exit Sub
    End If
    
    
    'Activo el libro que estoy por abrir
    ThisWorkbook.Activate
    
    
    Application.DisplayAlerts = False
    'Worksheets.Add
    'ActiveSheet.Name = "Errores"
    Application.DisplayAlerts = True
    'Set wsError = Worksheets("Errores")
    'Va la Hoja del Libro que se va a Abrir QUE SIEMPRE VA A SER EL DE HORAS EXTRAS
    Set wsContenido = wbContenido.Worksheets("Jur 2 Y 51 - Horas Extras 09-20")
    
    'va el nombre de la hoja del libro que ya tengo abierto que es el 120 del mensual
    Sheets("Hoja1").Select
    
    'Calculo el número de filas de la hoja actual del libro que ya tengo abierto del 120
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calculo el número de filas de la hoja de horas extras
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    'a la primer columna vacia del archivo que ya tengo abierto el 120 la nombro para ver cuales son las coincidencias con el archivo que voy a abrir de hs extras
    Cells(1, nColumnas + 1).Value = "IGUALES"
    
    
    For i = 2 To nFilas
        'Trato la hoja del archivo que ya tengo abierto el 120
        If Cells(i, 8).Value = "120" Then
            'Tomo el dni del archivo que ya tengo abierto el 120
            valorDoc = Cells(i, 5).Value
            'defino el rango de busqueda del archivo que voy a abrir el de hs extras concatenando D con nFilasCont para que me quede "D-nº=D12" y el numero defino del final del rango
            rangoTemp = "E2:E" & nFilasCont
            Set resultado = wsContenido.Range(rangoTemp).Find(What:=valorDoc, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False, _
                        SearchFormat:=False)
            'Si el resultado de la búsqueda no es vacío
            If Not resultado Is Nothing Then
                
                
                'Estoy en la fila del 120 correspondiente al dni encontrado en hs extras
                Cells(i, nColumnas + 1).Value = "COINCIDENCIA-DESCONTAR"
                
            Else
                Cells(i, nColumnas + 1).Value = "NO DESCONTAR"
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
       
                
                
                If Sheets("Hoja1").Cells(i, 16) = "COINCIDENCIA-DESCONTAR" Then
                    filaResultado = filaResultado + 1
                    Sheets("RESULTADO").Cells(filaResultado, 1).Value = 0
                    Sheets("RESULTADO").Cells(filaResultado, 2).Value = Sheets("Hoja1").Cells(i, 2).Value
                    Sheets("RESULTADO").Cells(filaResultado, 3).Value = Sheets("Hoja1").Cells(i, 3).Value
                    Sheets("RESULTADO").Cells(filaResultado, 4).Value = 0
                    Sheets("RESULTADO").Cells(filaResultado, 5).Value = Sheets("Hoja1").Cells(i, 5).Value
                    Sheets("RESULTADO").Cells(filaResultado, 6).Value = 0
                    Sheets("RESULTADO").Cells(filaResultado, 7).Value = Sheets("Hoja1").Cells(i, 7).Value
                    Sheets("RESULTADO").Cells(filaResultado, 8).Value = 120
                    Sheets("RESULTADO").Cells(filaResultado, 9).Value = 2
                    Sheets("RESULTADO").Cells(filaResultado, 10).Value = 25
                    Sheets("RESULTADO").Cells(filaResultado, 11).Value = Sheets("Hoja1").Cells(i, 12).Value
                    Sheets("RESULTADO").Cells(filaResultado, 12).Value = "92020"
               End If
    Next i
    
                
       
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub

End Sub

    
