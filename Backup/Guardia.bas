Attribute VB_Name = "Módulo1"
Sub Guardias()
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim i As Integer
    Dim filaError As Integer
    Dim fila275 As Integer
    Dim fila276 As Integer
    Dim fila277 As Integer
    Dim ultCuof As Integer
    Dim ultAnexo As Integer
    Dim ultDni As String
    Dim wsError As Excel.Worksheet
    Dim rechazosCuof As Integer, _
        horaNueva As Integer, _
        horaRechazada276 As Integer, _
        horaRechazada275 As Integer, _
        horaRechazada277 As Integer, _
        totalHora As Integer
    Dim bandera As Boolean
    
    
    'Agrego la nueva hoja
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Horas rechazadas"
    Application.DisplayAlerts = True
    Set wsError = Worksheets("Horas rechazadas")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count
    
    MsgBox "Debe estar ordenado por CUOF, ANEXO y DNI", , "Atención!!"
    'Se podría ver si se puede ordenar solo
    
    Range("1:1").Copy
    wsError.Range("1:1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    filaError = 2
    
    Cells(1, 9).Value = "HORAS"
    Cells(1, nColumnas + 1).Value = "HS_Rechazadas"
    
    Range("A1").Copy
    Range("M1").PasteSpecial Paste:=xlFormats
    Application.CutCopyMode = False
    
    i = 2
    ultCuof = Cells(i, 1).Value
    ultAnexo = Cells(i, 2).Value
    rechazosCuof = 0
    ultDni = "0"
    fila275 = 0
    fila276 = 0
    fila277 = 0
    Do While i <= nFilas
        If (ultCuof = Cells(i, 1).Value) And (ultAnexo = Cells(i, 2).Value) Then
            If ultDni <> Cells(i, 5).Value Then
                'Trato el documento anterior. Controlo fila y modifico si es necesario
                totalHora = 0
                horaRechazada276 = 0
                horaRechazada277 = 0
                horaRechazada275 = 0
                bandera = False
                If fila276 > 0 Then
                    totalHora = Cells(fila276, 9).Value
                    If totalHora > 300 Then
                        horaNueva = 300
                        horaRechazada276 = Cells(fila276, 9).Value - horaNueva
                        totalHora = 300
                        Cells(fila276, 9).Value = horaNueva
                        Cells(fila276, 9).Font.Bold = True
                        Cells(fila276, 9).Font.ColorIndex = 5
                        bandera = True
                        Cells(fila276, nColumnas + 1).Value = horaRechazada276
                        Cells(fila276, nColumnas + 1).Font.ColorIndex = 5
                    End If
                End If
                If fila275 > 0 Then
                    totalHora = totalHora + Cells(fila275, 9).Value
                    If totalHora > 300 Then
                        horaRechazada275 = totalHora - 300
                        horaNueva = Cells(fila275, 9).Value - horaRechazada275
                        totalHora = 300
                        Cells(fila275, 9).Value = horaNueva
                        Cells(fila275, 9).Font.Bold = True
                        Cells(fila275, 9).Font.ColorIndex = 5
                        bandera = True
                        Cells(fila275, nColumnas + 1).Value = horaRechazada275
                        Cells(fila275, nColumnas + 1).Font.ColorIndex = 5
                    End If
                End If
                If fila277 > 0 Then
                    totalHora = totalHora + Cells(fila277, 9).Value
                    If totalHora > 300 Then
                        horaRechazada277 = totalHora - 300
                        horaNueva = Cells(fila277, 9).Value - horaRechazada277
                        totalHora = 300
                        Cells(fila277, 9).Value = horaNueva
                        Cells(fila277, 9).Font.Bold = True
                        Cells(fila277, 9).Font.ColorIndex = 5
                        bandera = True
                        Cells(fila277, nColumnas + 1).Value = horaRechazada277
                        Cells(fila277, nColumnas + 1).Font.ColorIndex = 5
                    End If
                End If
                If bandera Then
                    rechazosCuof = rechazosCuof + 1
                    'Copia en la nueva hoja
                    wsError.Cells(filaError, 1).Value = Cells(i - 1, 1).Value
                    wsError.Cells(filaError, 2).Value = Cells(i - 1, 2).Value
                    wsError.Cells(filaError, 3).Value = Cells(i - 1, 3).Value
                    wsError.Cells(filaError, 4).Value = Cells(i - 1, 4).Value
                    wsError.Cells(filaError, 5).Value = Cells(i - 1, 5).Value
                    wsError.Cells(filaError, 6).Value = Cells(i - 1, 6).Value
                    wsError.Cells(filaError, 7).Value = Cells(i - 1, 7).Value
                    'wsError.Cells(filaError, 8).Value = Cells(i - 1, 8).Value
                    wsError.Cells(filaError, 9).Value = horaRechazada275
                    wsError.Cells(filaError, 10).Value = horaRechazada276
                    wsError.Cells(filaError, 11).Value = horaRechazada277
                    wsError.Cells(filaError, 12).Value = Cells(i - 1, 12).Value
                    filaError = filaError + 1
                End If
                ultDni = Cells(i, 5).Value
                fila275 = 0
                fila276 = 0
                fila277 = 0
            End If
            'Guardo las filas segun concepto
            If Cells(i, 8).Value = 276 Then
                fila276 = i
                Cells(i, 9).Value = Cells(i, 10).Value
            Else
                If Cells(i, 8).Value = 277 Then
                    fila277 = i
                    Cells(i, 9).Value = Cells(i, 11).Value
                Else
                    fila275 = i
                End If
            End If
            'Corrijo el valor de Anexo
            If Cells(i, 2).Value = 99 Then
                Cells(i, 2).Value = 0
            End If
            'Aumento el contador
            i = i + 1
        Else
            'Nuevo valor de Anexo
            ultCuof = Cells(i, 1).Value
            ultAnexo = Cells(i, 2).Value
            rechazosCuof = 0
            ultDni = "0"
            'No actualizo contador, así vuelve al Do y hace el proceso
        End If
    Loop
    'Eliminar columnas 10 y 11
    Columns(11).Delete
    Columns(10).Delete
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub


Sub Guardias_Extra()
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    'nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count
    
    i = 3
    cont = 1
    temp = Cells(2, 1).Value
    Do While i <= nFilas
        If temp = Cells(i, 1).Value Then
            cont = cont + 1
        Else
            If cont = 1 Then
                Cells(i - 1, 13).Value = "ÚNICO EN LA CUOF"
                Cells(i - 1, 13).Font.Bold = True
            End If
            cont = 1
            temp = Cells(i, 1).Value
        End If
        i = i + 1
    Loop
    If cont = 1 Then
        Cells(i - 1, 13).Value = "ÚNICO EN LA CUOF"
        Cells(i - 1, 13).Font.Bold = True
    End If
    
    Cells(1, 9).Value = "HS CPTO 275"
    Cells(1, 10).Value = "HS CPTO 276"
    Cells(1, 11).Value = "HS CPTO 277"
    Columns(8).Delete
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Suma_Horas()

    Dim suma275 As Double
    Dim suma276 As Double
    Dim suma277 As Double
    
     'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    'nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count
    
    suma275 = 0
    suma276 = 0
    suma277 = 0
    For i = 2 To nFilas
        If Cells(i, 8).Value = 275 Then
           suma275 = suma275 + Cells(i, 9).Value
        Else
            If Cells(i, 8).Value = 276 Then
                suma276 = suma276 + Cells(i, 9).Value
            Else
                suma277 = suma277 + Cells(i, 9).Value
            End If
        End If
    Next i
    Cells(nFilas + 1, 7).Value = suma275
    Cells(nFilas + 1, 8).Value = suma276
    Cells(nFilas + 1, 9).Value = suma277
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Tratar_500()
    Dim i As Long
    Dim rango As Range
    Dim rangoCont As Range
    Dim resultado As Range
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim nFilasCont As Double
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet, _
        wsError As Excel.Worksheet
    Dim temp As String
    Dim fila275 As Integer
    Dim fila276 As Integer
    Dim fila277 As Integer
    Dim horaNueva As Integer, _
        horaRechazada276 As Integer, _
        horaRechazada275 As Integer, _
        horaRechazada277 As Integer, _
        totalHora As Integer
    Dim bandera As Boolean
    Dim filaError As Integer
    
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
       
    'Agrego la nueva hoja
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Horas rechazadas"
    Application.DisplayAlerts = True
    Set wsError = Worksheets("Horas rechazadas")
    
    Set wsContenido = wbContenido.Worksheets("Hoja1")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count
    
    MsgBox "Debe estar ordenado por DNI, T.PROF(D) Y CoUC (276,275,277)", , "Atención!!"
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    Range("1:1").Copy
    wsError.Range("1:1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    nfilaError = 1
    
    Cells(1, 9).Value = "HORAS"
    Cells(1, nColumnas + 1).Value = "HS_Rechazadas"
    Cells(1, nColumnas + 2).Value = "Agentes 500hs"
    wsError.Cells(1, 13).Value = "OBSERVACIONES"
    
    Range("A1").Copy
    Range("M1").PasteSpecial Paste:=xlFormats
    Application.CutCopyMode = False
    
    For i = 2 To nFilasCont
        j = 0
        valorDoc = wsContenido.Cells(i, 2).Value
        'Busca en el otro archivo
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
            primerResultado = resultado.Address
            'Se obtiene el valor de j
            celdaDoc = resultado.Address
            tempDoc = ""
            For m = 1 To Len(celdaDoc)
                If IsNumeric(Mid(celdaDoc, m, 1)) Then
                    tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                End If
            Next m
            j = tempDoc
            
            'Empieza a tratar a la persona
            totalHora = 0
            Do While wsContenido.Cells(i, 2).Value = Cells(j, 5).Value
                horaRechazada = 0
                Cells(j, nColumnas + 2).Value = "Procesado"
                'No importa que oficina.. se toma la que está en resolución
                If Cells(j, 8).Value = 276 Then
                    If (totalHora + Cells(j, 10).Value) > 500 Then
                        'Rechazo horas
                        horaRechazada = (totalHora + Cells(j, 10).Value) - 500
                        Cells(j, 9).Value = Cells(j, 10).Value - horaRechazada
                        totalHora = 500
                        Cells(j, 9).Font.Bold = True
                        Cells(j, 9).Font.ColorIndex = 5
                        Cells(j, nColumnas + 1).Value = horaRechazada
                        Cells(j, nColumnas + 1).Font.ColorIndex = 5
                        
                        If wsError.Cells(nfilaError, 5).Value = wsContenido.Cells(i, 2).Value Then
                            wsError.Cells(nfilaError, 10).Value = wsError.Cells(nfilaError, 10).Value + horaRechazada
                        Else
                            'Agrego nuevo en ERROR
                            nfilaError = nfilaError + 1
                            wsError.Cells(nfilaError, 5).Value = Cells(j, 5).Value
                            wsError.Cells(nfilaError, 6).Value = Cells(j, 6).Value
                            wsError.Cells(nfilaError, 7).Value = Cells(j, 7).Value
                            wsError.Cells(nfilaError, 9).Value = 0
                            wsError.Cells(nfilaError, 10).Value = horaRechazada
                            wsError.Cells(nfilaError, 11).Value = 0
                            wsError.Cells(nfilaError, 12).Value = Cells(j, 12).Value
                        End If
                    Else
                        'Acepto todas las horas
                        Cells(j, 9).Value = Cells(j, 10).Value
                        totalHora = totalHora + Cells(j, 10).Value
                    End If
                Else
                    If Cells(j, 8).Value = 275 Then
                        If (totalHora + Cells(j, 9).Value) > 500 Then
                            'Rechazo horas
                            horaRechazada = (totalHora + Cells(j, 9).Value) - 500
                            Cells(j, 9).Value = Cells(j, 9).Value - horaRechazada
                            totalHora = 500
                            Cells(j, 9).Font.Bold = True
                            Cells(j, 9).Font.ColorIndex = 5
                            Cells(j, nColumnas + 1).Value = horaRechazada
                            Cells(j, nColumnas + 1).Font.ColorIndex = 5
                            
                            If wsError.Cells(nfilaError, 5).Value = wsContenido.Cells(j, 2).Value Then
                                wsError.Cells(nfilaError, 9).Value = wsError.Cells(nfilaError, 9).Value + horaRechazada
                            Else
                                'Agrego nuevo en ERROR
                                nfilaError = nfilaError + 1
                                wsError.Cells(nfilaError, 5).Value = Cells(j, 5).Value
                                wsError.Cells(nfilaError, 6).Value = Cells(j, 6).Value
                                wsError.Cells(nfilaError, 7).Value = Cells(j, 7).Value
                                wsError.Cells(nfilaError, 9).Value = horaRechazada
                                wsError.Cells(nfilaError, 10).Value = 0
                                wsError.Cells(nfilaError, 11).Value = 0
                                wsError.Cells(nfilaError, 12).Value = Cells(j, 12).Value
                            End If
                        Else
                            'Acepto todas las horas
                            totalHora = totalHora + Cells(j, 9).Value
                        End If
                    Else
                        If (totalHora + Cells(j, 11).Value) > 500 Then
                            'Rechazo horas
                            horaRechazada = (totalHora + Cells(j, 11).Value) - 500
                            Cells(j, 9).Value = Cells(j, 11).Value - horaRechazada
                            totalHora = 500
                            Cells(j, 9).Font.Bold = True
                            Cells(j, 9).Font.ColorIndex = 5
                            Cells(j, nColumnas + 1).Value = horaRechazada
                            Cells(j, nColumnas + 1).Font.ColorIndex = 5
                            
                            If wsError.Cells(nfilaError, 5).Value = wsContenido.Cells(i, 2).Value Then
                                wsError.Cells(nfilaError, 11).Value = wsError.Cells(nfilaError, 11).Value + horaRechazada
                            Else
                                'Agrego nuevo en ERROR
                                nfilaError = nfilaError + 1
                                wsError.Cells(nfilaError, 5).Value = Cells(j, 5).Value
                                wsError.Cells(nfilaError, 6).Value = Cells(j, 6).Value
                                wsError.Cells(nfilaError, 7).Value = Cells(j, 7).Value
                                wsError.Cells(nfilaError, 9).Value = 0
                                wsError.Cells(nfilaError, 10).Value = 0
                                wsError.Cells(nfilaError, 11).Value = horaRechazada
                                wsError.Cells(nfilaError, 12).Value = Cells(j, 12).Value
                            End If
                        Else
                            'Acepto todas las horas
                            Cells(j, 9).Value = Cells(j, 11).Value
                            totalHora = totalHora + Cells(j, 11).Value
                        End If
                    End If
                End If
                j = j + 1
            Loop
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub

Sub Tratar_300()
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim i As Integer
    Dim nfilaError As Integer
    Dim ultDni As String
    Dim wsError As Excel.Worksheet
    Dim rechazosCuof As Integer, _
        horaNueva As Integer, _
        horaRechazada As Integer, _
        totalHora As Integer
    Dim bandera As Boolean
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count
    
    Set wsError = Worksheets("Horas rechazadas")
    Set rangoError = wsError.UsedRange
    nfilaError = rangoError.Rows.Count
    
    MsgBox "Debe estar ordenado por DNI, T.PROF(D) Y CoUC (276,275,277)", , "Atención!!"
    
    nfilaError = nfilaError + 1
    
    i = 2
    totalHora = 0
    ultDni = Cells(i, 5).Value
    horaRechazada = 0
    Do While i <= nFilas
        If Cells(i, nColumnas) <> "Procesado" Then
            If ultDni = Cells(i, 5).Value Then
                If Cells(i, 8).Value = 276 Then
                    If (totalHora + Cells(i, 10).Value) > 300 Then
                        'Rechazo horas
                        horaRechazada = (totalHora + Cells(i, 10).Value) - 300
                        Cells(i, 9).Value = Cells(i, 10).Value - horaRechazada
                        totalHora = 300
                        Cells(i, 9).Font.Bold = True
                        Cells(i, 9).Font.ColorIndex = 5
                        Cells(i, nColumnas - 1).Value = horaRechazada
                        Cells(i, nColumnas - 1).Font.ColorIndex = 5
                        
                        If wsError.Cells(nfilaError, 5).Value = ultDni Then
                            wsError.Cells(nfilaError, 10).Value = wsError.Cells(nfilaError, 10).Value + horaRechazada
                        Else
                            'Agrego nuevo en ERROR
                            nfilaError = nfilaError + 1
                            wsError.Cells(nfilaError, 5).Value = Cells(i, 5).Value
                            wsError.Cells(nfilaError, 6).Value = Cells(i, 6).Value
                            wsError.Cells(nfilaError, 7).Value = Cells(i, 7).Value
                            wsError.Cells(nfilaError, 9).Value = 0
                            wsError.Cells(nfilaError, 10).Value = horaRechazada
                            wsError.Cells(nfilaError, 11).Value = 0
                            wsError.Cells(nfilaError, 12).Value = Cells(i, 12).Value
                        End If
                    Else
                        'Acepto todas las horas
                        Cells(i, 9).Value = Cells(i, 10).Value
                        totalHora = totalHora + Cells(i, 10).Value
                    End If
                Else
                    If Cells(i, 8).Value = 275 Then
                        If (totalHora + Cells(i, 9).Value) > 300 Then
                            'Rechazo horas
                            horaRechazada = (totalHora + Cells(i, 9).Value) - 300
                            Cells(i, 9).Value = Cells(i, 9).Value - horaRechazada
                            totalHora = 300
                            Cells(i, 9).Font.Bold = True
                            Cells(i, 9).Font.ColorIndex = 5
                            Cells(i, nColumnas - 1).Value = horaRechazada
                            Cells(i, nColumnas - 1).Font.ColorIndex = 5
                            
                            If wsError.Cells(nfilaError, 5).Value = ultDni Then
                                wsError.Cells(nfilaError, 9).Value = wsError.Cells(nfilaError, 9).Value + horaRechazada
                            Else
                                'Agrego nuevo en ERROR
                                nfilaError = nfilaError + 1
                                wsError.Cells(nfilaError, 5).Value = Cells(i, 5).Value
                                wsError.Cells(nfilaError, 6).Value = Cells(i, 6).Value
                                wsError.Cells(nfilaError, 7).Value = Cells(i, 7).Value
                                wsError.Cells(nfilaError, 9).Value = horaRechazada
                                wsError.Cells(nfilaError, 10).Value = 0
                                wsError.Cells(nfilaError, 11).Value = 0
                                wsError.Cells(nfilaError, 12).Value = Cells(i, 12).Value
                            End If
                        Else
                            'Acepto todas las horas
                            totalHora = totalHora + Cells(i, 9).Value
                        End If
                    Else
                        If (totalHora + Cells(i, 11).Value) > 300 Then
                            'Rechazo horas
                            horaRechazada = (totalHora + Cells(i, 11).Value) - 300
                            Cells(i, 9).Value = Cells(i, 11).Value - horaRechazada
                            totalHora = 300
                            Cells(i, 9).Font.Bold = True
                            Cells(i, 9).Font.ColorIndex = 5
                            Cells(i, nColumnas - 1).Value = horaRechazada
                            Cells(i, nColumnas - 1).Font.ColorIndex = 5
                            
                            If wsError.Cells(nfilaError, 5).Value = ultDni Then
                                wsError.Cells(nfilaError, 11).Value = wsError.Cells(nfilaError, 11).Value + horaRechazada
                            Else
                                'Agrego nuevo en ERROR
                                nfilaError = nfilaError + 1
                                wsError.Cells(nfilaError, 5).Value = Cells(i, 5).Value
                                wsError.Cells(nfilaError, 6).Value = Cells(i, 6).Value
                                wsError.Cells(nfilaError, 7).Value = Cells(i, 7).Value
                                wsError.Cells(nfilaError, 9).Value = 0
                                wsError.Cells(nfilaError, 10).Value = 0
                                wsError.Cells(nfilaError, 11).Value = horaRechazada
                                wsError.Cells(nfilaError, 12).Value = Cells(i, 12).Value
                            End If
                        Else
                            'Acepto todas las horas
                            Cells(i, 9).Value = Cells(i, 11).Value
                            totalHora = totalHora + Cells(i, 11).Value
                        End If
                    End If
                End If
            Else
                totalHora = 0
                ultDni = Cells(i, 5).Value
                horaRechazada = 0
                i = i - 1
            End If
        End If
        
        i = i + 1
    Loop
    'Eliminar columnas 10 y 11
    Columns(11).Delete
    Columns(10).Delete
    'Eliminar columnas de ERROR
    wsError.Columns(12).Delete
    wsError.Columns(8).Delete
    wsError.Columns(4).Delete
    wsError.Columns(3).Delete
    wsError.Columns(2).Delete
    wsError.Columns(1).Delete
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Control_Pago()
    Dim i As Long
    Dim valorDoc As String
    Dim rango As Range
    Dim rangoCont As Range
    Dim resultado As Range
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim nColumnasCont As Double
    Dim nFilasCont As Double
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet
    Dim temp As String
    Dim bandera As Boolean
    Dim filaError As Integer
    
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
       
        
    Set wsContenido = wbContenido.Worksheets("Hoja1")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    nColumnasCont = rangoCont.Columns.Count
    
    For i = 2 To nFilas
        j = 0
        valorDoc = Cells(i, 4).Value
        'Busca en el otro archivo
        rangoTemp = "L2:L" & nFilasCont
        Set resultado = wsContenido.Range(rangoTemp).Find(What:=valorDoc, _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False, _
                    SearchFormat:=False)
        
        'Si el resultado de la búsqueda no es vacío
        If Not resultado Is Nothing Then
            primerResultado = resultado.Address
            'Se obtiene el valor de j
            celdaDoc = resultado.Address
            tempDoc = ""
            For m = 1 To Len(celdaDoc)
                If IsNumeric(Mid(celdaDoc, m, 1)) Then
                    tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                End If
            Next m
            j = tempDoc
            
            bandera = True
            tempAlerta = 0
            Do While wsContenido.Cells(j, 12).Value = valorDoc
                If Cells(i, 7).Value = wsContenido.Cells(j, 4).Value And wsContenido.Cells(j, nColumnasCont + 1).Value = "" Then
                    If Cells(i, 3).Value = wsContenido.Cells(j, 20).Value Then
                        bandera = False
                        wsContenido.Cells(j, nColumnasCont + 1).Value = "Procesado"
                        If Cells(i, 8).Value = wsContenido.Cells(j, 18).Value Then
                            Cells(i, nColumnas + 1).Value = "Ok"
                        Else
                            Cells(i, nColumnas + 1).Value = "Cargado:"
                            Cells(i, nColumnas + 2).Value = wsContenido.Cells(j, 18).Value
                        End If
                    Else
                        tempAlerta = j
                    End If
                End If
                j = j + 1
            Loop
            If bandera Then
                Cells(i, nColumnas + 1).Value = "No liquidado"
            End If
            If tempAlerta <> 0 Then
                If wsContenido.Cells(tempAlerta, nColumnasCont + 1) = "" Then
                    Cells(i, nColumnas + 1).Value = "Otro Cuof"
                    Cells(i, nColumnas + 2).Value = wsContenido.Cells(tempAlerta, 18).Value
                    wsContenido.Cells(tempAlerta, nColumnasCont + 1).Value = "Procesado"
                End If
            End If
        Else
            Cells(i, nColumnas + 2).Value = "No liquidado. No encontró DNI"
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub

Sub Control_Pago_3()
    Dim i As Long
    Dim valorDoc As String
    Dim rango As Range
    Dim rangoCont As Range
    Dim resultado As Range
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim nColumnasCont As Double
    Dim nFilasCont As Double
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet
    Dim importe As Double
    Dim horas As Double
    Dim importeCalc As Double
    Dim horasCalc As Double
    
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
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    Set wsContenido = wbContenido.Worksheets("Hoja1")
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    nColumnasCont = rangoCont.Columns.Count
    
    For i = 2 To nFilas
        j = 0
        importe = 0
        horas = 0
        importeCalc = Cells(i, nColumnas).Value
        horasCalc = Cells(i, 9).Value
        Do While (Cells(i, 5).Value = Cells(i + 1, 5).Value And Cells(i, 8).Value = Cells(i + 1, 8).Value)
            i = i + 1
            importeCalc = importeCalc + Cells(i, nColumnas).Value
            horasCalc = horasCalc + Cells(i, 9).Value
        Loop
        valorDoc = Cells(i, 5).Value
        
        Cells(i, nColumnas + 1).Value = horasCalc
        Cells(i, nColumnas + 2).Value = importeCalc
        
        'Busca en el otro archivo
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
            primerResultado = resultado.Address
            'Se obtiene el valor de j
            celdaDoc = resultado.Address
            tempDoc = ""
            For m = 1 To Len(celdaDoc)
                If IsNumeric(Mid(celdaDoc, m, 1)) Then
                    tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                End If
            Next m
            j = tempDoc
            
            Do While wsContenido.Cells(j - 1, 5).Value = valorDoc
                j = j - 1
            Loop
            Do While wsContenido.Cells(j, 5).Value = valorDoc
                If wsContenido.Cells(j, 11).Value = Cells(i, 8).Value Then
                    If wsContenido.Cells(j, 14).Value = 2 Then
                        importe = importe - wsContenido.Cells(j, 17).Value
                        horas = horas - wsContenido.Cells(j, 15).Value
                    Else
                        importe = importe + wsContenido.Cells(j, 17).Value
                        horas = horas + wsContenido.Cells(j, 15).Value
                    End If
                End If
                j = j + 1
            Loop
            Cells(i, nColumnas + 3).Value = horas
            Cells(i, nColumnas + 4).Value = importe
        Else
            Cells(i, nColumnas + 5).Value = "Error. No encontró DNI"
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub

Sub Control_Pago_2()
    Dim i As Long
    Dim valorDoc As String
    Dim rango As Range
    Dim rangoCont As Range
    Dim resultado As Range
    Dim nFilasPago As Long
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim nColumnasCont As Double
    Dim nFilasCont As Double
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet, _
        wsPago As Excel.Worksheet
    Dim importe As Double
    
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

    'Agrego la nueva hoja
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Montos a Pagar"
    Application.DisplayAlerts = True
    Set wsPago = Worksheets("Montos a Pagar")
    Set wsContenido = wbContenido.Worksheets("Hoja1")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    nColumnasCont = rangoCont.Columns.Count
    
    wsPago.Cells(1, 1).Value = Cells(1, 5).Value
    wsPago.Cells(1, 2).Value = Cells(1, 6).Value
    wsPago.Cells(1, 3).Value = Cells(1, 7).Value
    wsPago.Cells(1, 4).Value = "Importe"
    
    nFilasPago = 1
    
    For i = 2 To nFilas
        If wsPago.Cells(nFilasPago, 1).Value = Cells(i, 5).Value Then
            wsPago.Cells(nFilasPago, 4).Value = wsPago.Cells(nFilasPago, 4).Value + Cells(i, nColumnas).Value
        Else
            nFilasPago = nFilasPago + 1
            wsPago.Cells(nFilasPago, 1).Value = Cells(i, 5).Value
            wsPago.Cells(nFilasPago, 2).Value = Cells(i, 6).Value
            wsPago.Cells(nFilasPago, 3).Value = Cells(i, 7).Value
            wsPago.Cells(nFilasPago, 4).Value = Cells(i, nColumnas).Value
        End If
    Next i
    
    For i = 2 To nFilasPago
        j = 0
        importe = 0
        valorDoc = wsPago.Cells(i, 1).Value
        'Busca en el otro archivo
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
            primerResultado = resultado.Address
            'Se obtiene el valor de j
            celdaDoc = resultado.Address
            tempDoc = ""
            For m = 1 To Len(celdaDoc)
                If IsNumeric(Mid(celdaDoc, m, 1)) Then
                    tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                End If
            Next m
            j = tempDoc
            
            Do While wsContenido.Cells(j - 1, 5).Value = valorDoc
                j = j - 1
            Loop
            Do While wsContenido.Cells(j, 5).Value = valorDoc
                If wsContenido.Cells(j, 14).Value = 2 Then
                    importe = importe - wsContenido.Cells(j, 17).Value
                Else
                    importe = importe + wsContenido.Cells(j, 17).Value
                End If
                j = j + 1
            Loop
            wsPago.Cells(i, 5).Value = importe
        Else
            wsPago.Cells(i, 6).Value = "Error. No encontró DNI"
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub

Sub Controlar_Doc()
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim i As Integer
    
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    For i = 2 To (nFilas - 1)
        j = j + 1
        Do While j < nFilas
            If Cells(j, 5).Value = Cells(i, 5).Value And Cells(j, 15).Value = Cells(i, 15).Value Then
                If Cells(j, 7).Value = Cells(i, 7).Value Then
                    Cells(j, nColumnas + 1).Value = "Repetido"
                    Cells(i, nColumnas + 1).Value = "Repetido"
                    temp1 = i & ":" & i
                    Rows(temp1).Interior.Color = RGB(178, 255, 102)
                    temp2 = j & ":" & j
                    Rows(temp2).Interior.Color = RGB(178, 255, 102)
                End If
            Else
                j = nFilas
            End If
            j = j + 1
        Loop
    Next i
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Formato_Recepción()
    Dim i As Long
    Dim valorDoc As String
    Dim rango As Range
    Dim nFilas As Double
    Dim nColumnas As Double
    
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count

    i = 2
    Do While i < nFilas + 1
        Rows(i + 1).Insert
        Rows(i + 2).Insert
        Rows(i + 3).Insert
        
        For j = 1 To 7
            Cells(i + 1, j).Value = Cells(i, j).Value
            Cells(i + 2, j).Value = Cells(i, j).Value
            Cells(i + 3, j).Value = Cells(i, j).Value
        Next j
        
        Cells(i + 1, 8).Value = 275
        Cells(i + 1, 9).Value = Cells(i, 9).Value
        Cells(i + 2, 8).Value = 276
        Cells(i + 2, 10).Value = Cells(i, 10).Value
        Cells(i + 3, 8).Value = 277
        Cells(i + 3, 11).Value = Cells(i, 11).Value
        
        For j = 12 To 15
            Cells(i + 1, j).Value = Cells(i, j).Value
            Cells(i + 2, j).Value = Cells(i, j).Value
            Cells(i + 3, j).Value = Cells(i, j).Value
        Next j
        
        Rows(i).Delete
        i = i + 3
        nFilas = nFilas + 2
    Loop

    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Formato_Recepción_2()
    Dim i As Long
    Dim valorDoc As String
    Dim rango As Range
    Dim nFilas As Double
    Dim nColumnas As Double
    
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count

    For i = 2 To nFilas
        If Cells(i, 8).Value = 275 Then
            Cells(i, 10).Value = ""
            Cells(i, 11).Value = ""
        Else
            If Cells(i, 8).Value = 276 Then
                Cells(i, 9).Value = ""
                Cells(i, 11).Value = ""
            Else
                Cells(i, 9).Value = ""
                Cells(i, 10).Value = ""
            End If
        End If
    Next

    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Eliminar_Ceros()
    Dim i As Long
    Dim rango As Range
    Dim nFilas As Double
    Dim nColumnas As Double
    
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count

    i = 2
    Do While i < nFilas + 1
        'If (Cells(i, 9).Value = 0 Or Cells(i, 10).Value = 0 Or Cells(i, 11).Value = 0 Or (Cells(i, 9).Value = "" And Cells(i, 10).Value = "" And Cells(i, 11).Value = "")) Then
        If Cells(i, 9).Value = 0 And Cells(i, 10).Value = 0 And Cells(i, 11).Value = 0 Then
            Rows(i).Delete
            i = i - 1
            nFilas = nFilas - 1
        End If
        i = i + 1
    Loop
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Cargar_TProf()
    Dim i As Long
    Dim valorDoc As String
    Dim rango As Range
    Dim rangoCont As Range
    Dim resultado As Range
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim nColumnasCont As Double
    Dim nFilasCont As Double
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet
    Dim temp As String
    
    
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
       
    Set wsContenido = wbContenido.Worksheets("Hoja1")
    
    'Regresa el control a la hoja de origen
    Sheets(1).Select
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count

    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    nColumnasCont = rangoCont.Columns.Count
    
    For i = 2 To nFilas
        j = 0
        'Si cpto es 277 no lo proceso porque el pago es igual para todos
        If Not (Cells(i, 8) = 277) Then
            valorDoc = Cells(i, 5).Value
            'Busca en el otro archivo
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
                primerResultado = resultado.Address
                'Se obtiene el valor de j
                celdaDoc = resultado.Address
                tempDoc = ""
                For m = 1 To Len(celdaDoc)
                    If IsNumeric(Mid(celdaDoc, m, 1)) Then
                        tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                    End If
                Next m
                j = tempDoc
                
                If wsContenido.Cells(j, nColumnasCont + 1).Value = "" Then
                    wsContenido.Cells(j, nColumnasCont + 1).Value = Cells(i, 7).Value
                Else
                    If Not (wsContenido.Cells(j, nColumnasCont + 1).Value = Cells(i, 7).Value) Then
                        wsContenido.Cells(j, nColumnasCont + 2).Value = Cells(i, 7).Value
                    End If
                End If
            Else
                wsContenido.Cells(nFilasCont + 1, 1).Value = Cells(i, 5).Value
                wsContenido.Cells(nFilasCont + 1, 2).Value = Cells(i, 6).Value
                wsContenido.Cells(nFilasCont + 1, nColumnasCont + 1).Value = Cells(i, 7).Value
                nFilasCont = nFilasCont + 1
            End If
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub

Sub Comparar_TProf()
    Dim i As Long
    Dim j As Long
    Dim rango As Range
    Dim nFilas As Double
    Dim nColumnas As Double
    
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count

    For i = 2 To nFilas
        ult = ""
        
        For j = 4 To nColumnas
            If ult = "" Then
                ult = Cells(i, j).Value
            Else
                If Cells(i, j).Value <> "" Then
                    If ult <> Cells(i, j).Value Then
                        Cells(i, nColumnas + 1).Value = "Controlar"
                    End If
                End If
            End If
        Next j
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Calcular_Monto()
    Dim i As Long
    Dim j As Long
    Dim rango As Range
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim importe As Long
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count

    For i = 2 To nFilas
        importe = 0
        If Cells(i, 8).Value = 276 Then
            If Cells(i, 7).Value = "A" Then
                importe = Cells(i, 9).Value * 150
            Else
                If Cells(i, 7).Value = "B" Then
                    importe = Cells(i, 9).Value * 140
                Else
                    importe = Cells(i, 9).Value * 85
                End If
            End If
        Else
            If Cells(i, 8).Value = 275 Then
                If Cells(i, 7).Value = "A" Then
                    importe = Cells(i, 9).Value * 100
                Else
                    If Cells(i, 7).Value = "B" Then
                        importe = Cells(i, 9).Value * 90
                    Else
                        importe = Cells(i, 9).Value * 70
                    End If
                End If
            Else
                If Cells(i, 8).Value = 277 Then
                    importe = Cells(i, 9).Value * 40
                Else
                    Cells(i, nColumnas + 2).Value = "Controlar"
                End If
            End If
        End If
        Cells(i, nColumnas + 1).Value = importe
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Calcular_Monto_Nuevo()
    Dim i As Long
    Dim j As Long
    Dim rango As Range
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim importe As Long
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count

    For i = 2 To nFilas
        importe = 0
        If Cells(i, 8).Value = 276 Then
            If Cells(i, 7).Value = "D" Then
                importe = Cells(i, 9).Value * 325
            Else
                If Cells(i, 7).Value = "A" Then
                    importe = Cells(i, 9).Value * 225
                Else
                    If Cells(i, 7).Value = "B" Then
                        importe = Cells(i, 9).Value * 210
                    Else
                        importe = Cells(i, 9).Value * 127
                    End If
                End If
            End If
        Else
            If Cells(i, 8).Value = 275 Then
                If Cells(i, 7).Value = "D" Then
                    importe = Cells(i, 9).Value * 250
                Else
                    If Cells(i, 7).Value = "A" Then
                        importe = Cells(i, 9).Value * 150
                    Else
                        If Cells(i, 7).Value = "B" Then
                            importe = Cells(i, 9).Value * 135
                        Else
                            importe = Cells(i, 9).Value * 105
                        End If
                    End If
                End If
            Else
                If Cells(i, 8).Value = 277 Then
                    importe = Cells(i, 9).Value * 60
                Else
                    Cells(i, nColumnas + 2).Value = "Controlar"
                End If
            End If
        End If
        Cells(i, nColumnas + 1).Value = importe
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub


Sub Cargar_Desc_Faltantes()
    Dim i As Long
    Dim valorDoc As String
    Dim rango As Range
    Dim rangoCont As Range
    Dim resultado As Range
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim nColumnasCont As Double
    Dim nFilasCont As Double
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet
    Dim temp As String
    Dim bandera As Boolean
    
    
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
       
    Set wsContenido = wbContenido.Worksheets("Hoja1")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count

    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    nColumnasCont = rangoCont.Columns.Count
    
    filas = nFilas
    For i = 2 To nFilasCont
        j = 0
        bandera = True
        valorDoc = wsContenido.Cells(i, 5).Value
        'Busca en el otro archivo
        rangoTemp = "E2:E" & filas
        Set resultado = Range(rangoTemp).Find(What:=valorDoc, _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False, _
                    SearchFormat:=False)
        'Si el resultado de la búsqueda no es vacío
        If Not resultado Is Nothing Then
            primerResultado = resultado.Address
            'Se obtiene el valor de j
            celdaDoc = resultado.Address
            tempDoc = ""
            For m = 1 To Len(celdaDoc)
                If IsNumeric(Mid(celdaDoc, m, 1)) Then
                    tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                End If
            Next m
            j = tempDoc
            
            Do While Cells(j - 1, 5).Value = valorDoc
                j = j - 1
            Loop
            Do While Cells(j, 5).Value = valorDoc
                If wsContenido.Cells(i, 11).Value = Cells(j, 8).Value Then
                    bandera = False
                End If
                j = j + 1
            Loop
            
            If bandera Then
                nFilas = nFilas + 1
                j = j - 1
                Cells(nFilas, 1).Value = Cells(j, 1).Value
                Cells(nFilas, 2).Value = Cells(j, 2).Value
                Cells(nFilas, 3).Value = Cells(j, 3).Value
                Cells(nFilas, 4).Value = Cells(j, 4).Value
                Cells(nFilas, 5).Value = Cells(j, 5).Value
                Cells(nFilas, 6).Value = Cells(j, 6).Value
                Cells(nFilas, 7).Value = Cells(j, 7).Value
                Cells(nFilas, 8).Value = wsContenido.Cells(i, 11).Value
                Cells(nFilas, 9).Value = wsContenido.Cells(i, 15).Value
                Cells(nFilas, 10).Value = "MAL Pagado"
                Cells(nFilas, 11).Value = wsContenido.Cells(i, 14).Value
            End If
        Else
            nFilas = nFilas + 1
            Cells(nFilas, 5).Value = wsContenido.Cells(i, 5).Value
            Cells(nFilas, 6).Value = wsContenido.Cells(i, 7).Value
            Cells(nFilas, 8).Value = wsContenido.Cells(i, 11).Value
            Cells(nFilas, 9).Value = wsContenido.Cells(i, 15).Value
            Cells(nFilas, 10).Value = "No encontrado"
            Cells(nFilas, 11).Value = wsContenido.Cells(i, 14).Value
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


Sub Resumen_1()
    Dim i As Long
    Dim valorDoc As String
    Dim rango As Range
    Dim rangoCont As Range
    Dim resultado As Range
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim nColumnasCont As Double
    Dim nFilasCont As Double
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet
    Dim temp As String
    Dim bandera As Boolean
    
    
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
       
    Set wsContenido = wbContenido.Worksheets("Hoja1")
    
    'Regresa el control a la hoja de origen
    Sheets("Resumen").Select
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count

    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    nColumnasCont = rangoCont.Columns.Count
    
    For i = 2 To nFilasCont
        j = 0
        bandera = True
        valorDoc = wsContenido.Cells(i, 5).Value
        valor275 = 0
        valor276 = 0
        valor277 = 0
        i = i - 1
        Do While wsContenido.Cells(i + 1, 5).Value = valorDoc
            i = i + 1
            If wsContenido.Cells(i, 11).Value = 275 Then
                If wsContenido.Cells(i, 14).Value = 2 Then
                    valor275 = valor275 - wsContenido.Cells(i, 15).Value
                Else
                    valor275 = valor275 + wsContenido.Cells(i, 15).Value
                End If
            Else
                If wsContenido.Cells(i, 11).Value = 276 Then
                    If wsContenido.Cells(i, 14).Value = 2 Then
                        valor276 = valor276 - wsContenido.Cells(i, 15).Value
                    Else
                        valor276 = valor276 + wsContenido.Cells(i, 15).Value
                    End If
                Else
                    If wsContenido.Cells(i, 11).Value = 277 Then
                        If wsContenido.Cells(i, 14).Value = 2 Then
                            valor277 = valor277 - wsContenido.Cells(i, 15).Value
                        Else
                            valor277 = valor277 + wsContenido.Cells(i, 15).Value
                        End If
                    End If
                End If
            End If
        Loop
        
        'Busca en el otro archivo
        rangoTemp = "A2:A" & nFilas
        Set resultado = Range(rangoTemp).Find(What:=valorDoc, _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False, _
                    SearchFormat:=False)
        'Si el resultado de la búsqueda no es vacío
        If Not resultado Is Nothing Then
            primerResultado = resultado.Address
            'Se obtiene el valor de j
            celdaDoc = resultado.Address
            tempDoc = ""
            For m = 1 To Len(celdaDoc)
                If IsNumeric(Mid(celdaDoc, m, 1)) Then
                    tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                End If
            Next m
            j = tempDoc
            
            If Cells(j, 6).Value <> valor275 And valor275 > 0 Then
                If Cells(j, 6).Value = "" Then
                    Cells(j, 6).Value = valor275
                Else
                    Cells(j, 6).Value = Cells(j, 6).Value + " - " + valor275
                End If
            End If
            
            If Cells(j, 7).Value <> valor276 And valor276 > 0 Then
                If Cells(j, 7).Value = "" Then
                    Cells(j, 7).Value = valor276
                Else
                    Cells(j, 7).Value = Cells(j, 7).Value + " - " + valor276
                End If
            End If
            
            If Cells(j, 8).Value <> valor277 And valor277 > 0 Then
                If Cells(j, 8).Value = "" Then
                    Cells(j, 8).Value = valor277
                Else
                    Cells(j, 8).Value = Cells(j, 8).Value + " - " + valor277
                End If
            End If
        Else
            nFilas = nFilas + 1
            Cells(nFilas, 1).Value = wsContenido.Cells(i, 5).Value
            Cells(nFilas, 2).Value = wsContenido.Cells(i, 7).Value
            If valor275 > 0 Then
                Cells(nFilas, 6).Value = valor275
            End If
            If valor276 > 0 Then
                Cells(nFilas, 7).Value = valor276
            End If
            If valor277 > 0 Then
                Cells(nFilas, 8).Value = valor277
            End If
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub

Sub Resumen_2()
    Dim i As Long
    Dim valorDoc As String
    Dim rango As Range
    Dim rangoResumen As Range
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim nFilasRes As Double
    Dim nFilasCarga As Double
    Dim wsResultado As Excel.Worksheet, _
        wsCarga As Excel.Worksheet
    Dim bandera As Boolean
    Dim temp As Integer
    
    
    'Activar este libro
    ThisWorkbook.Activate
       
    Set wsResultado = Worksheets("Resumen")
    
    'Regresa el control a la hoja de origen
    Sheets("Descuentos").Select
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count
    
    Set rangoResumen = wsResultado.UsedRange
    nFilasRes = rangoResumen.Rows.Count
    
    For i = 2 To nFilas
        temp = 0
        If Cells(i, 8).Value = 275 Then
            temp = 15
        Else
            If Cells(i, 8) = 276 Then
                temp = 16
            Else
                If Cells(i, 8) = 277 Then
                    temp = 17
                End If
            End If
        End If
    
        valorDoc = Cells(i, 5).Value
        'Busca en el otro archivo
        rangoTemp = "A2:A" & nFilasRes
        Set resultado = wsResultado.Range(rangoTemp).Find(What:=valorDoc, _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False, _
                    SearchFormat:=False)
        'Si el resultado de la búsqueda no es vacío
        If Not resultado Is Nothing Then
            primerResultado = resultado.Address
            'Se obtiene el valor de j
            celdaDoc = resultado.Address
            tempDoc = ""
            For m = 1 To Len(celdaDoc)
                If IsNumeric(Mid(celdaDoc, m, 1)) Then
                    tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                End If
            Next m
            j = tempDoc
            
            wsResultado.Cells(j, temp).Value = Cells(i, 9).Value
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub


Sub Control_Prof()
    Dim i As Long
    Dim valorDoc As String
    Dim rango As Range
    Dim rangoResumen As Range
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim nFilasRes As Double
    Dim nFilasCarga As Double
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet
    Dim bandera As Boolean
    Dim temp As Integer
    
    
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
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    Set wsContenido = wbContenido.Worksheets("Hoja1")
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    nColumnasCont = rangoCont.Columns.Count
    
    For i = 2 To nFilas
        valorDoc = Cells(i, 5).Value
        'Busca en el otro archivo
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
            primerResultado = resultado.Address
            'Se obtiene el valor de j
            celdaDoc = resultado.Address
            tempDoc = ""
            For m = 1 To Len(celdaDoc)
                If IsNumeric(Mid(celdaDoc, m, 1)) Then
                    tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                End If
            Next m
            j = tempDoc
            
            Cells(i, nColumnas + 1).Value = wsContenido.Cells(j, 6).Value
            
            If Cells(i, nColumnas + 1).Value <> Cells(i, 7).Value Then
                Cells(i, nColumnas + 2).Value = "Controlar"
                If Cells(i, nColumnas + 1).Value = "A" And Cells(i, 7).Value = "D" Then
                    Cells(i, nColumnas + 2).Value = ""
                End If
            End If
            If wsContenido.Cells(j, 3).Value <> "" Then
                Cells(i, nColumnas + 2).Value = "Controlar. C.Obra"
            End If
        Else
            Cells(i, nColumnas + 2).Value = "Controlar"
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub

Sub Control_Prof_D()
    Dim i As Long
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim bandera As Boolean
    Dim importe64 As Integer
    Dim importe105 As Integer
    Dim importe106 As Integer
    Dim importe126 As Integer
    Dim importe241 As Integer
    Dim importe247 As Integer
    Dim importe254 As Integer
    Dim importe255 As Integer
    Dim wsError As Excel.Worksheet
    
    Sheets("Hoja1").Select
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count
    
    Set wsError = Worksheets("Horas rechazadas")
    Set rangoError = wsError.UsedRange
    nfilaError = rangoError.Rows.Count
    
    importe105 = 0
    importe106 = 0
    importe126 = 0
    importe241 = 0
    importe247 = 0
    importe254 = 0
    importe255 = 0
    For i = 2 To nFilas
        bandera = False
        If Cells(i, 7).Value = "D" Then
            If Cells(i, 1).Value = 64 Then
                temp = 1680
                If (importe64 + Cells(i, 9).Value) > temp Then
                    Cells(i, nColumnas - 1).Value = importe64 + Cells(i, 9).Value - temp
                    Cells(i, 9).Value = temp - importe64
                    importe64 = temp
                    bandera = True
                Else
                    importe64 = importe64 + Cells(i, 9).Value
                End If
            Else
                If Cells(i, 1).Value = 105 Then
                    temp = 4500
                    If (importe105 + Cells(i, 9).Value) > temp Then
                        Cells(i, nColumnas - 1).Value = importe105 + Cells(i, 9).Value - temp
                        Cells(i, 9).Value = temp - importe105
                        importe105 = temp
                        bandera = True
                    Else
                        importe105 = importe105 + Cells(i, 9).Value
                    End If
                Else
                    If Cells(i, 1).Value = 106 Then
                        temp = 10000
                        If (importe106 + Cells(i, 9).Value) > temp Then
                            Cells(i, nColumnas - 1).Value = importe106 + Cells(i, 9).Value - temp
                            Cells(i, 9).Value = temp - importe106
                            importe106 = temp
                            bandera = True
                        Else
                            importe106 = importe106 + Cells(i, 9).Value
                        End If
                    Else
                        If Cells(i, 1).Value = 126 Then
                            temp = 1680
                            If (importe126 + Cells(i, 9).Value) > temp Then
                                Cells(i, nColumnas - 1).Value = importe126 + Cells(i, 9).Value - temp
                                Cells(i, 9).Value = temp - importe126
                                importe126 = temp
                                bandera = True
                            Else
                                importe126 = importe126 + Cells(i, 9).Value
                            End If
                        Else
                            If Cells(i, 1).Value = 241 Then
                                temp = 7000
                                If (importe241 + Cells(i, 9).Value) > temp Then
                                    Cells(i, nColumnas - 1).Value = importe241 + Cells(i, 9).Value - temp
                                    Cells(i, 9).Value = temp - importe241
                                    importe241 = temp
                                    bandera = True
                                Else
                                    importe241 = importe241 + Cells(i, 9).Value
                                End If
                            Else
                                If Cells(i, 1).Value = 247 Then
                                    temp = 2800
                                    If (importe247 + Cells(i, 9).Value) > temp Then
                                        Cells(i, nColumnas - 1).Value = importe247 + Cells(i, 9).Value - temp
                                        Cells(i, 9).Value = temp - importe247
                                        importe247 = temp
                                        bandera = True
                                    Else
                                        importe247 = importe247 + Cells(i, 9).Value
                                    End If
                                Else
                                    If Cells(i, 1).Value = 254 Then
                                        temp = 1680
                                        If (importe254 + Cells(i, 9).Value) > temp Then
                                            Cells(i, nColumnas - 1).Value = importe254 + Cells(i, 9).Value - temp
                                            Cells(i, 9).Value = temp - importe254
                                            importe254 = temp
                                            bandera = True
                                        Else
                                            importe254 = importe254 + Cells(i, 9).Value
                                        End If
                                    Else
                                        If Cells(i, 1).Value = 255 Then
                                            temp = 1680
                                            If (importe255 + Cells(i, 9).Value) > temp Then
                                                Cells(i, nColumnas - 1).Value = importe255 + Cells(i, 9).Value - temp
                                                Cells(i, 9).Value = temp - importe255
                                                importe255 = temp
                                                bandera = True
                                            Else
                                                importe255 = importe255 + Cells(i, 9).Value
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            If bandera Then
                nfilaError = nfilaError + 1
                wsError.Cells(nfilaError, 1).Value = Cells(i, 5).Value
                wsError.Cells(nfilaError, 2).Value = Cells(i, 6).Value
                wsError.Cells(nfilaError, 3).Value = "D"
                wsError.Cells(nfilaError, 7).Value = "Horas Criticas"
                If Cells(i, 8).Value = 275 Then
                    wsError.Cells(nfilaError, 4).Value = Cells(i, nColumnas - 1).Value
                Else
                    If Cells(i, 8).Value = 276 Then
                        wsError.Cells(nfilaError, 5).Value = Cells(i, nColumnas - 1).Value
                    Else
                        wsError.Cells(nfilaError, 6).Value = Cells(i, nColumnas - 1).Value
                    End If
                End If
            End If
        End If
    Next i
    
    'Informe
    nfilaError = nfilaError + 2
    wsError.Cells(nfilaError, 1).Value = "TOTAL HORAS CRITICAS"
    nfilaError = nfilaError + 1
    wsError.Cells(nfilaError, 1).Value = "CUOF 64"
    wsError.Cells(nfilaError, 2).Value = importe64
    nfilaError = nfilaError + 1
    wsError.Cells(nfilaError, 1).Value = "CUOF 105"
    wsError.Cells(nfilaError, 2).Value = importe105
    nfilaError = nfilaError + 1
    wsError.Cells(nfilaError, 1).Value = "CUOF 106"
    wsError.Cells(nfilaError, 2).Value = importe106
    nfilaError = nfilaError + 1
    wsError.Cells(nfilaError, 1).Value = "CUOF 126"
    wsError.Cells(nfilaError, 2).Value = importe126
    nfilaError = nfilaError + 1
    wsError.Cells(nfilaError, 1).Value = "CUOF 241"
    wsError.Cells(nfilaError, 2).Value = importe241
    nfilaError = nfilaError + 1
    wsError.Cells(nfilaError, 1).Value = "CUOF 247"
    wsError.Cells(nfilaError, 2).Value = importe247
    nfilaError = nfilaError + 1
    wsError.Cells(nfilaError, 1).Value = "CUOF 254"
    wsError.Cells(nfilaError, 2).Value = importe254
    nfilaError = nfilaError + 1
    wsError.Cells(nfilaError, 1).Value = "CUOF 255"
    wsError.Cells(nfilaError, 2).Value = importe255
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub

Sub Control_Horas_Cuof()
    Dim i As Long
    Dim valorDoc As String
    Dim rango As Range
    Dim rangoResumen As Range
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim nFilasRes As Long
    Dim nFilasCarga As Double
    Dim wbContenido As Workbook, _
        wsResultado As Excel.Worksheet, _
        wsContenido As Excel.Worksheet
    Dim bandera As Boolean
    Dim temp As Integer
    
    
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
       
    MsgBox "Debe estar ordenado por Cuof + Anexo.", , "Atención!!"
    
    
    'Agrego la nueva hoja
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Control Horas"
    Application.DisplayAlerts = True
    Set wsResultado = Worksheets("Control Horas")
    
    Set wsContenido = wbContenido.Worksheets("Hoja1")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    nColumnasCont = rangoCont.Columns.Count
    
    
    nFilasRes = 1
    wsResultado.Cells(nFilasRes, 1).Value = "Cuof"
    wsResultado.Cells(nFilasRes, 2).Value = "Anexo"
    wsResultado.Cells(nFilasRes, 3).Value = "Act. Críticas"
    wsResultado.Cells(nFilasRes, 4).Value = "Activas"
    wsResultado.Cells(nFilasRes, 5).Value = "Pasivas"
    wsResultado.Cells(nFilasRes, 6).Value = "MAX Act. Críticas"
    wsResultado.Cells(nFilasRes, 7).Value = "MAX Activas"
    wsResultado.Cells(nFilasRes, 8).Value = "MAX Pasivas"
    wsResultado.Cells(nFilasRes, 9).Value = "Observación"
    
    nFilasRes = 2
    wsResultado.Cells(nFilasRes, 1).Value = Cells(2, 1).Value
    wsResultado.Cells(nFilasRes, 2).Value = Cells(2, 2).Value
    wsResultado.Cells(nFilasRes, 3).Value = 0
    wsResultado.Cells(nFilasRes, 4).Value = 0
    wsResultado.Cells(nFilasRes, 5).Value = 0
    
    'VER QUE ESTÉN BIEN UBICADAS LAS COLUMNAS 276-275-277
    For i = 2 To nFilas
        If Cells(i, 1).Value = wsResultado.Cells(nFilasRes, 1).Value And Cells(i, 2).Value = wsResultado.Cells(nFilasRes, 2).Value Then
            If Cells(i, 7).Value = "D" Then
                If Cells(i, 8).Value = "275" Then
                    wsResultado.Cells(nFilasRes, 3).Value = wsResultado.Cells(nFilasRes, 3).Value + Cells(i, 9).Value
                Else
                    'wsResultado.Cells(nFilasRes, 3).Value = wsResultado.Cells(nFilasRes, 3).Value + Cells(i, 9).Value
                    wsResultado.Cells(nFilasRes, 3).Value = wsResultado.Cells(nFilasRes, 3).Value + Cells(i, 10).Value
                End If
            Else
                If Cells(i, 8).Value = "277" Then
                    'wsResultado.Cells(nFilasRes, 5).Value = wsResultado.Cells(nFilasRes, 5).Value + Cells(i, 9).Value
                    wsResultado.Cells(nFilasRes, 5).Value = wsResultado.Cells(nFilasRes, 5).Value + Cells(i, 11).Value
                Else
                    If Cells(i, 8).Value = "275" Then
                        wsResultado.Cells(nFilasRes, 4).Value = wsResultado.Cells(nFilasRes, 4).Value + Cells(i, 9).Value
                    Else
                        'wsResultado.Cells(nFilasRes, 4).Value = wsResultado.Cells(nFilasRes, 4).Value + Cells(i, 9).Value
                        wsResultado.Cells(nFilasRes, 4).Value = wsResultado.Cells(nFilasRes, 4).Value + Cells(i, 10).Value
                    End If
                End If
            End If
        Else
            nFilasRes = nFilasRes + 1
            wsResultado.Cells(nFilasRes, 1).Value = Cells(i, 1).Value
            wsResultado.Cells(nFilasRes, 2).Value = Cells(i, 2).Value
            wsResultado.Cells(nFilasRes, 3).Value = 0
            wsResultado.Cells(nFilasRes, 4).Value = 0
            wsResultado.Cells(nFilasRes, 5).Value = 0
            
            i = i - 1
        End If
    Next i
    
    For i = 2 To nFilasRes
        bandera = True
        For j = 3 To 115
            If wsContenido.Cells(j, 2).Value = wsResultado.Cells(i, 1).Value Then
                If wsContenido.Cells(j, 3).Value = wsResultado.Cells(i, 2).Value Then
                    wsResultado.Cells(i, 6).Value = wsContenido.Cells(j, 8).Value
                    wsResultado.Cells(i, 7).Value = wsContenido.Cells(j, 9).Value
                    wsResultado.Cells(i, 8).Value = wsContenido.Cells(j, 10).Value
                    
                    If wsResultado.Cells(i, 3).Value > wsContenido.Cells(j, 8).Value Then
                        wsResultado.Cells(i, 9).Value = "Controlar"
                    Else
                        If wsResultado.Cells(i, 4).Value > wsContenido.Cells(j, 9).Value Then
                            wsResultado.Cells(i, 9).Value = "Controlar"
                        Else
                            If wsResultado.Cells(i, 5).Value > wsContenido.Cells(j, 10).Value Then
                                wsResultado.Cells(i, 9).Value = "Controlar"
                            End If
                        End If
                    End If
                    bandera = False
                    j = 115
                End If
            End If
        Next j
        If bandera Then
            wsResultado.Cells(i, 9).Value = "No encontró Jur+Anexo"
        End If
        For m = 1 To 3
            dif = wsResultado.Cells(i, 2 + j).Value - wsResultado.Cells(i, 5 + j).Value
            If dif > 0 Then
                wsResultado.Cells(i, 8 + j).Value = dif
            End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'. Valor i: " & i, , "Error"
End Sub
