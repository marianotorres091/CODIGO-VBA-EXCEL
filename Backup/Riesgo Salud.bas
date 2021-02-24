Attribute VB_Name = "Módulo11"
Sub Filtrar_Resolucion()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i As Long
    Dim filaCopia As Long
    Dim nFilasError As Integer
    Dim columnaCopia As Long
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet


    'Indicar el libro de excel CONTENIDO y control de errores
    contenido = InputBox("Ingrese el nombre del archivo:", "Abrir", "Archivo.xlsx")
    If contenido <> "" Then
        On Error GoTo ControlErrorOpen
        Set wbContenido = Workbooks.Open(ActiveWorkbook.Path & "\" & contenido)
    Else
        Exit Sub
    End If
    
    'Activar este libro
    ThisWorkbook.Activate
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Errores"
    Application.DisplayAlerts = True
    Set wsError = Worksheets("Errores")
    Set wsContenido = wbContenido.Worksheets("Hoja1")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    'Sheets(1).Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    Range("1:1").Copy
    wsError.Range("1:1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    nFilasError = 2
    
    Cells(1, nColumnas + 1).Value = "Resolución"
    Cells(1, nColumnas + 2).Value = "CUOF seg Res."
    
    For i = 2 To nFilasCont
        If wsContenido.Cells(i, 5).Value <> "" Then
            'Buscar en el otro archivo y marcar
            valorDoc = wsContenido.Cells(i, 3).Value
            rangoTemp = "E2:E" & nFilas
            Set resultado = Range(rangoTemp).Find(What:=valorDoc, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False, _
                        SearchFormat:=False)
            'Si el resultado de la búsqueda no es vacío
            If Not resultado Is Nothing Then
                'Se obtiene el valor de j
                celdaDoc = resultado.Address
                tempDoc = ""
                For m = 1 To Len(celdaDoc)
                    If IsNumeric(Mid(celdaDoc, m, 1)) Then
                        tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                    End If
                Next m
                j = tempDoc
                
                'Estoy en la fila correspondiente al doc
                Cells(j, nColumnas + 1).Value = wsContenido.Cells(i, 5).Value
                Cells(j, nColumnas + 2).Value = wsContenido.Cells(i, 6).Value
            Else
                'Copiar en ERRORES
                temp = i & ":" & i
                wsContenido.Range(temp).Copy
                temp = nFilasError & ":" & nFilasError
                wsError.Range(temp).PasteSpecial xlPasteAll
                Application.CutCopyMode = False
                nFilasError = nFilasError + 1
            End If
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


Sub Procesar_Resolucion()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i As Long
    Dim filaCont As Long
    Dim filaAgente As Long
    Dim nFilasNuevo As Long
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet, _
        wsNuevo As Excel.Worksheet, _
        wsAgente As Worksheet
    Dim cuof As Integer
    Dim anexo As Integer
    Dim totalCobrado As Long
    Dim totalNoCobrado As Long
    
    
    'Indicar el libro de excel CONTENIDO y control de errores
    contenido = InputBox("Ingrese el nombre del archivo:", "Abrir", "Archivo.xlsx")
    If contenido <> "" Then
        On Error GoTo ControlErrorOpen
        Set wbContenido = Workbooks.Open(ActiveWorkbook.Path & "\" & contenido)
    Else
        Exit Sub
    End If

    'Activar este libro
    ThisWorkbook.Activate
    
    'ActiveSheet.Name = "Totales"
    
    Application.DisplayAlerts = False
    wbContenido.Worksheets.Add
    wbContenido.ActiveSheet.Name = "Agentes sin Cobrar"
    Worksheets.Add
    ActiveSheet.Name = "Agentes"
    Application.DisplayAlerts = True
    Set wsAgente = Worksheets("Agentes")
    Set wsNuevo = wbContenido.Worksheets("Agentes sin Cobrar")
    Set wsContenido = wbContenido.Worksheets("Hoja1")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    wsContenido.Cells(1, 1).Value = "Porcentaje de Personas a Cobrar:"
    wsContenido.Cells(1, 3).Value = 0.2
    
    wsContenido.Cells(2, 1).Value = "Cuof"
    wsContenido.Cells(2, 2).Value = "Anexo"
    wsContenido.Cells(2, 3).Value = "Cant. Cobran R.Salud"
    wsContenido.Cells(2, 4).Value = "Cant. NO Cobran R.Salud"
    wsContenido.Cells(2, 5).Value = "Cant. Agregados para Cobrar"
    
    filaCont = 3
    
    Range("1:1").Copy
    wsNuevo.Range("1:1").PasteSpecial xlPasteAll
    wsAgente.Range("1:1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    nFilasNuevo = 2
    filaAgente = 2
    
    MsgBox "Debe estar ordenado por CUOF y ANEXO.", , "Atención"
    
    totalCobrado = 0
    totalNoCobrado = 0
    'cuof = Cells(2, 16).Value
    'anexo = Cells(2, 17).Value
    cuof = 0
    anexo = 0
    
    For i = 2 To nFilas
        If Cells(i, nColumnas).Value <> "" Then
            'Copio en una nueva hoja
            temp = i & ":" & i
            Range(temp).Copy
            temp = filaAgente & ":" & filaAgente
            wsAgente.Range(temp).PasteSpecial xlPasteAll
            Application.CutCopyMode = False
            filaAgente = filaAgente + 1
            
            nFilasNuevo = nFilasNuevo + 1
            'Tratar registro
            If cuof = Cells(i, 16).Value And anexo = Cells(i, 17).Value Then
                'Acumulo en donde corresponda
                If Cells(i, 32).Value > 0 Then
                    totalCobrado = totalCobrado + 1
                Else
                    totalNoCobrado = totalNoCobrado + 1
                    
                    temp = i & ":" & i
                    Range(temp).Copy
                    temp = nFilasNuevo & ":" & nFilasNuevo
                    wsNuevo.Range(temp).PasteSpecial xlPasteAll
                    Application.CutCopyMode = False
                    nFilasNuevo = nFilasNuevo + 1
                End If
            Else
                'Copio totales y reinicio
                wsContenido.Cells(filaCont, 3).Value = totalCobrado
                wsContenido.Cells(filaCont, 4).Value = totalNoCobrado
                wsContenido.Cells(filaCont, 5).Value = "=D" & filaCont & "*C1"
                
                filaCont = filaCont + 1
                cuof = Cells(i, 16).Value
                anexo = Cells(i, 17).Value
                wsContenido.Cells(filaCont, 1).Value = Cells(i, 16).Value
                wsContenido.Cells(filaCont, 2).Value = Cells(i, 17).Value
                totalCobrado = 0
                totalNoCobrado = 0
                
                'Acumulo en donde corresponda
                If Cells(i, 32).Value > 0 Then
                    totalCobrado = totalCobrado + 1
                Else
                    totalNoCobrado = totalNoCobrado + 1
                    
                    temp = i & ":" & i
                    Range(temp).Copy
                    temp = nFilasNuevo & ":" & nFilasNuevo
                    wsNuevo.Range(temp).PasteSpecial xlPasteAll
                    Application.CutCopyMode = False
                    nFilasNuevo = nFilasNuevo + 1
                End If
            End If
        End If
    Next i
    wsContenido.Cells(filaCont, 3).Value = totalCobrado
    wsContenido.Cells(filaCont, 4).Value = totalNoCobrado
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub

Sub Copiar_Porcentaje()
    Dim wsContenido As Excel.Worksheet
    Dim nFilasCont As Double
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim nFilasCont As Long
    Dim i As Long
    Dim filaCont As Long
    Dim filaDest As Long
    Dim nfilaCont As Long
    Dim cuof As Integer
    Dim anexo As Integer
    Dim totalCobrar As Long

    'Indica el libro de excel donde se guarda
    destino = InputBox("Ingrese el nombre del archivo donde se desea guardar:", "Abrir", "Archivo.xlsx")
    If destino <> "" Then
        On Error GoTo ControlErrorOpen
        Set wbDestino = Workbooks.Open(ActiveWorkbook.Path & "\" & destino)
    Else
        Exit Sub
    End If
    
    'Activar este libro
    ThisWorkbook.Activate
    Set wsDestino = wbDestino.Worksheets(1)
    
    'Regresa el control a la hoja de origen
    Sheets("Agentes sin Cobrar").Select
    
    Set wsContenido = Worksheets("Totales")
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    MsgBox "Debe estar ordenado por CUOF, ANEXO y de Menor a Mayor los importes.", , "Atención"
    
    filaCont = 3
    cuof = wsContenido.Cells(filaCont, 1).Value
    anexo = wsContenido.Cells(filaCont, 2).Value
    totalCobrar = wsContenido.Cells(filaCont, 5).Value
    
    
    wsDestino.Cells(1, 1).Value = "PtaId"
    wsDestino.Cells(1, 2).Value = "JurId"
    wsDestino.Cells(1, 3).Value = "EscId"
    wsDestino.Cells(1, 4).Value = "Pref"
    wsDestino.Cells(1, 5).Value = "Doc"
    wsDestino.Cells(1, 6).Value = "Digito"
    wsDestino.Cells(1, 7).Value = "Nombres"
    wsDestino.Cells(1, 8).Value = "Couc"
    wsDestino.Cells(1, 9).Value = "Reajuste"
    wsDestino.Cells(1, 10).Value = "Unidades"
    wsDestino.Cells(1, 11).Value = "Importe"
    wsDestino.Cells(1, 12).Value = "Vto"
    wsDestino.Range("1:1").Font.Bold = True
    wsDestino.Range("1:1").HorizontalAlignment = xlCenter
    
    filaDest = 2
    
    For i = 2 To nFilas
        'SE VIENE LA PAPA
        If cuof = Cells(i, 16).Value And anexo = Cells(i, 17).Value Then
            If totalCobrar > 0 Then
                'COPIO EL REGISTRO CON FORMATO DE CARGA A UNA NUEVA HOJA
                wsDestino.Cells(filaDest, 1).Value = 0
                wsDestino.Cells(filaDest, 2).Value = "JurId"
                wsDestino.Cells(filaDest, 3).Value = "EscId"
                wsDestino.Cells(filaDest, 4).Value = 0
                wsDestino.Cells(filaDest, 5).Value = Cells(i, 5).Value
                wsDestino.Cells(filaDest, 6).Value = 0
                wsDestino.Cells(filaDest, 7).Value = Cells(i, 7).Value
                wsDestino.Cells(filaDest, 8).Value = "Couc"
                wsDestino.Cells(filaDest, 9).Value = "Reajuste"
                wsDestino.Cells(filaDest, 10).Value = "Unidades"
                wsDestino.Cells(filaDest, 11).Value = Cells(i, 32).Value
                wsDestino.Cells(filaDest, 12).Value = "Vto"
                filaDest = filaDest + 1
                totalCobrar = totalCobrar - 1
            End If
        Else
            filaCont = filaCont + 1
            cuof = wsContenido.Cells(filaCont, 1).Value
            anexo = wsContenido.Cells(filaCont, 2).Value
            totalCobrar = wsContenido.Cells(filaCont, 5).Value
            If totalCobrar > 0 Then
                'COPIO EL REGISTRO CON FORMATO DE CARGA A UNA NUEVA HOJA
                wsDestino.Cells(filaDest, 1).Value = 0
                wsDestino.Cells(filaDest, 2).Value = "JurId"
                wsDestino.Cells(filaDest, 3).Value = "EscId"
                wsDestino.Cells(filaDest, 4).Value = 0
                wsDestino.Cells(filaDest, 5).Value = Cells(i, 5).Value
                wsDestino.Cells(filaDest, 6).Value = 0
                wsDestino.Cells(filaDest, 7).Value = Cells(i, 7).Value
                wsDestino.Cells(filaDest, 8).Value = "Couc"
                wsDestino.Cells(filaDest, 9).Value = "Reajuste"
                wsDestino.Cells(filaDest, 10).Value = "Unidades"
                wsDestino.Cells(filaDest, 11).Value = Cells(i, 32).Value
                wsDestino.Cells(filaDest, 12).Value = "Vto"
                filaDest = filaDest + 1
                totalCobrar = totalCobrar - 1
            End If
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


Sub Agentes_NoCobran()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i As Long
    Dim filaCont As Long
    Dim filaAgente As Long
    Dim nFilasNuevo As Long
    Dim wsNuevo As Excel.Worksheet, _
        wsContenido As Excel.Worksheet
    Dim cuof As Integer
    Dim anexo As Integer
    Dim totalCobrado As Long
    Dim totalNoCobrado As Long
    
    MsgBox "Debe estar ordenado por DNI y Concepto.", , "Atención!!"
    
    'Activar este libro
    ThisWorkbook.Activate
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Agentes sin Cobrar"
    Application.DisplayAlerts = True
    Set wsNuevo = Worksheets("Agentes sin Cobrar")
    Set wsContenido = Worksheets("Tabla")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    filaCont = 1
    wsNuevo.Cells(filaCont, 1).Value = "DNI"
    wsNuevo.Cells(filaCont, 2).Value = "NOMBRE"
    wsNuevo.Cells(filaCont, 3).Value = "CEIC"
    wsNuevo.Cells(filaCont, 4).Value = "246-R.SALUD"
    filaCont = 2
    
    i = 2
    Do
        If Cells(i, 23).Value = 1 Then
            apartado = ""
            If Cells(i, 4).Value = 1 And Cells(i + 1, 4).Value = 1 Then
                apartado = Cells(i, 24).Value
                grupo = Cells(i, 26).Value
            Else
                If Cells(i, 4).Value = 1 And Cells(i + 1, 4).Value = 100 And Cells(i + 2, 4).Value = 1 Then
                    i = i + 1
                    apartado = Cells(i, 24).Value
                    grupo = Cells(i, 26).Value
                Else
                    If Cells(i + 1, 4).Value = 246 Then
                        i = i + 1
                    Else
                        i = i + 2
                    End If
                End If
            End If
            
            If apartado <> "" Then
                For j = 3 To 42
                    If wsContenido.Cells(j, 2).Value = apartado And wsContenido.Cells(j, 3).Value = grupo Then
                        m = j
                        j = 42
                    End If
                Next j
            
                wsNuevo.Cells(filaCont, 1).Value = Cells(i, 12).Value
                wsNuevo.Cells(filaCont, 2).Value = Cells(i, 14).Value
                wsNuevo.Cells(filaCont, 3).Value = Cells(i, 15).Value
                wsNuevo.Cells(filaCont, 4).Value = wsContenido.Cells(m, 12).Value
                filaCont = filaCont + 1
            End If
        End If
        i = i + 1
    Loop While i < nFilas + 1
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


Sub Comparar_Arch()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i As Long
    Dim filaCopia As Long
    Dim nFilasError As Integer
    Dim columnaCopia As Long
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet


    'Indicar el libro de excel CONTENIDO y control de errores
    contenido = InputBox("Ingrese el nombre del archivo:", "Abrir", "Archivo.xlsx")
    If contenido <> "" Then
        On Error GoTo ControlErrorOpen
        Set wbContenido = Workbooks.Open(ActiveWorkbook.Path & "\" & contenido)
    Else
        Exit Sub
    End If
    
    'Activar este libro
    ThisWorkbook.Activate
    
    Set wsContenido = wbContenido.Worksheets("Agentes sin Cobrar")
    
    Sheets("RiesgoSalud 06-2018").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    Cells(1, nColumnas + 1).Value = "Correcto"
    Cells(1, nColumnas + 2).Value = "Importe"
    Cells(1, nColumnas + 3).Value = "Diferencia"
    
    For i = 2 To nFilas
        If Cells(i, 28).Value <> "Auditado" Then
            'Buscar en el otro archivo y marcar
            valorDoc = Cells(i, 4).Value
            rangoTemp = "A2:A" & nFilasCont
            Set resultado = wsContenido.Range(rangoTemp).Find(What:=valorDoc, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False, _
                        SearchFormat:=False)
            'Si el resultado de la búsqueda no es vacío
            If Not resultado Is Nothing Then
                'Se obtiene el valor de j
                celdaDoc = resultado.Address
                tempDoc = ""
                For m = 1 To Len(celdaDoc)
                    If IsNumeric(Mid(celdaDoc, m, 1)) Then
                        tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                    End If
                Next m
                j = tempDoc
                
                'Estoy en la fila correspondiente al doc
                Cells(i, nColumnas + 1).Value = "SI"
                Cells(i, nColumnas + 2).Value = wsContenido.Cells(j, 4).Value
                Cells(i, nColumnas + 3).Value = Cells(i, 27).Value - wsContenido.Cells(j, 4).Value
            Else
                Cells(i, nColumnas + 1).Value = "NO"
            End If
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


Sub Obtener_Carga()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim tempFecha As Date
    Dim i As Long
    Dim filaResultado As Long
    Dim filaCopia As Long
    Dim columnaCopia As Long
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet, _
        wsNuevo As Excel.Worksheet


    'Indicar el libro de excel CONTENIDO y control de errores
    contenido = InputBox("Ingrese el nombre del archivo:", "Abrir", "Archivo.xlsx")
    If contenido <> "" Then
        On Error GoTo ControlErrorOpen
        Set wbContenido = Workbooks.Open(ActiveWorkbook.Path & "\" & contenido)
    Else
        Exit Sub
    End If
    
    'Activar este libro
    ThisWorkbook.Activate
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "RESULTADO"
    Application.DisplayAlerts = True
    Set wsNuevo = Worksheets("RESULTADO")
    
    Set wsContenido = wbContenido.Worksheets(1)
    
    Sheets("Año 2015").Select
    
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    filaResultado = 1
    wsNuevo.Cells(filaResultado, 1).Value = "PtaId"
    wsNuevo.Cells(filaResultado, 2).Value = "JurId"
    wsNuevo.Cells(filaResultado, 3).Value = "EscId"
    wsNuevo.Cells(filaResultado, 4).Value = "Pref"
    wsNuevo.Cells(filaResultado, 5).Value = "Doc"
    wsNuevo.Cells(filaResultado, 6).Value = "Digito"
    wsNuevo.Cells(filaResultado, 7).Value = "Nombres"
    wsNuevo.Cells(filaResultado, 8).Value = "Couc"
    wsNuevo.Cells(filaResultado, 9).Value = "Reajuste"
    wsNuevo.Cells(filaResultado, 10).Value = "Unidades"
    wsNuevo.Cells(filaResultado, 11).Value = "Importe"
    wsNuevo.Cells(filaResultado, 12).Value = "Vto"
        
    
    For i = 2 To nFilas
        If Cells(i, 8).Value = 7 Then
            'Buscar en el otro archivo y marcar
            valorDoc = Cells(i, 4).Value
            rangoTemp = "H2:H" & nFilasCont
            Set resultado = wsContenido.Range(rangoTemp).Find(What:=valorDoc, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False, _
                        SearchFormat:=False)
            'Si el resultado de la búsqueda no es vacío
            If Not resultado Is Nothing Then
                'Se obtiene el valor de j
                celdaDoc = resultado.Address
                tempDoc = ""
                For m = 1 To Len(celdaDoc)
                    If IsNumeric(Mid(celdaDoc, m, 1)) Then
                        tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                    End If
                Next m
                j = tempDoc
                
                'Estoy en la fila correspondiente al doc
                bandera = True
                Do
                    tempFecha = wsContenido.Cells(j, 17).Value
                    If Year(tempFecha) = Cells(i, 7).Value And Month(tempFecha) = Cells(i, 8).Value Then
                        bandera = False
                        Cells(i, nColumnas + 1).Value = wsContenido.Cells(j, 14).Value
                    End If
                    j = j + 1
                Loop While valorDoc = wsContenido.Cells(j, 8).Value
                
                If bandera Then
                    filaResultado = filaResultado + 1
                    wsNuevo.Cells(filaResultado, 1).Value = 0
                    wsNuevo.Cells(filaResultado, 2).Value = Cells(i, 1).Value
                    wsNuevo.Cells(filaResultado, 3).Value = 2
                    wsNuevo.Cells(filaResultado, 4).Value = 0
                    wsNuevo.Cells(filaResultado, 5).Value = Cells(i, 4).Value
                    wsNuevo.Cells(filaResultado, 6).Value = 0
                    wsNuevo.Cells(filaResultado, 7).Value = Cells(i, 6).Value
                    wsNuevo.Cells(filaResultado, 8).Value = 246
                    wsNuevo.Cells(filaResultado, 9).Value = 1
                    wsNuevo.Cells(filaResultado, 10).Value = 0
                    wsNuevo.Cells(filaResultado, 11).Value = Cells(i, 21).Value
                    wsNuevo.Cells(filaResultado, 12).Value = Cells(i, 8).Value & Cells(i, 7).Value
                End If
            Else
                filaResultado = filaResultado + 1
                wsNuevo.Cells(filaResultado, 1).Value = 0
                wsNuevo.Cells(filaResultado, 2).Value = Cells(i, 1).Value
                wsNuevo.Cells(filaResultado, 3).Value = 2
                wsNuevo.Cells(filaResultado, 4).Value = 0
                wsNuevo.Cells(filaResultado, 5).Value = Cells(i, 4).Value
                wsNuevo.Cells(filaResultado, 6).Value = 0
                wsNuevo.Cells(filaResultado, 7).Value = Cells(i, 6).Value
                wsNuevo.Cells(filaResultado, 8).Value = 246
                wsNuevo.Cells(filaResultado, 9).Value = 1
                wsNuevo.Cells(filaResultado, 10).Value = 0
                wsNuevo.Cells(filaResultado, 11).Value = Cells(i, 21).Value
                wsNuevo.Cells(filaResultado, 12).Value = Cells(i, 8).Value & Cells(i, 7).Value
            End If
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub

