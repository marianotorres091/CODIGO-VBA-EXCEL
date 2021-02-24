Attribute VB_Name = "Module1"
Sub Control_Personas()
    Dim i As Long
    Dim rango As Range
    Dim rangoCont As Range
    Dim resultado As Range
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim nFilasCont As Double
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet, _
        wsResultado As Excel.Worksheet
    Dim temp As String
    Dim bandera As Boolean
    Dim filaResultado As Long
    Dim dni As String
    Dim fecha As Date
    
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
    ActiveSheet.Name = "C_Personas"
    Application.DisplayAlerts = True
    Set wsResultado = Worksheets("C_Personas")
    
    Set wsContenido = wbContenido.Worksheets("Hoja1")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count
    
    MsgBox "Debe estar ordenado por DNI + VTO", , "Atención!!"
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    filaResultado = 1
    wsResultado.Cells(filaResultado, 1).Value = "DNI"
    wsResultado.Cells(filaResultado, 2).Value = "Nombre"
    wsResultado.Cells(filaResultado, 3).Value = "Mes"
    wsResultado.Cells(filaResultado, 4).Value = "Año"
    wsResultado.Cells(filaResultado, 5).Value = "Hs_276_Crit"
    wsResultado.Cells(filaResultado, 6).Value = "Hs_275_Crit"
    wsResultado.Cells(filaResultado, 7).Value = "Hs_276"
    wsResultado.Cells(filaResultado, 8).Value = "Hs_275"
    wsResultado.Cells(filaResultado, 9).Value = "Hs_277"
    wsResultado.Cells(filaResultado, 10).Value = "Total Hs"
    wsResultado.Cells(filaResultado, 11).Value = "Observación"
    wsResultado.Cells(filaResultado, 12).Value = "Horas excedidas"
    
    dni = Cells(2, 12).Value
    ultMes = 0
    
    For i = 2 To nFilas
        fecha = Cells(i, 16).Value
        anio = Year(fecha)
        mes = Month(fecha)
    
        'faltaría comparar año, pero en este caso no es necesario
        If dni = Cells(i, 12).Value And ultMes = mes Then
            If Cells(i, 4) = "277" Then
                If Cells(i, 18) <> 0 Then
                    If Cells(i, 6) = 2 Then
                        wsResultado.Cells(filaResultado, 9).Value = wsResultado.Cells(filaResultado, 9).Value - Cells(i, 18).Value
                        wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value - Cells(i, 18).Value
                    Else
                        wsResultado.Cells(filaResultado, 9).Value = wsResultado.Cells(filaResultado, 9).Value + Cells(i, 18).Value
                        wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value + Cells(i, 18).Value
                    End If
                Else
                    If anio > 2017 And mes > 3 Then
                        hora = Cells(i, 7).Value \ 40
                    Else
                        hora = Cells(i, 7).Value \ 60
                    End If
                    If Cells(i, 6) = 2 Then
                        wsResultado.Cells(filaResultado, 9).Value = wsResultado.Cells(filaResultado, 9).Value - hora
                        wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value - hora
                    Else
                        wsResultado.Cells(filaResultado, 9).Value = wsResultado.Cells(filaResultado, 9).Value + hora
                        wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value + hora
                    End If
                End If
            Else
                If Cells(i, 4) = "276" Then
                    If Cells(i, 18).Value <> 0 Then
                        'ver si es critico
                        If Cells(i, 7).Value / Cells(i, 18).Value = 325 Then
                            If Cells(i, 6) = 2 Then
                                wsResultado.Cells(filaResultado, 5).Value = wsResultado.Cells(filaResultado, 5).Value - Cells(i, 18).Value
                                wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value - Cells(i, 18).Value
                            Else
                                wsResultado.Cells(filaResultado, 5).Value = wsResultado.Cells(filaResultado, 5).Value + Cells(i, 18).Value
                                wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value + Cells(i, 18).Value
                            End If
                        Else
                            If Cells(i, 6) = 2 Then
                                wsResultado.Cells(filaResultado, 7).Value = wsResultado.Cells(filaResultado, 7).Value - Cells(i, 18).Value
                                wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value - Cells(i, 18).Value
                            Else
                                wsResultado.Cells(filaResultado, 7).Value = wsResultado.Cells(filaResultado, 7).Value + Cells(i, 18).Value
                                wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value + Cells(i, 18).Value
                            End If
                        End If
                    Else
                        'suma todo a No critico
                        If anio > 2017 And mes > 3 Then
                            hora = Cells(i, 7).Value \ 150
                        Else
                            hora = Cells(i, 7).Value \ 225
                        End If
                        If Cells(i, 6) = 2 Then
                            wsResultado.Cells(filaResultado, 7).Value = wsResultado.Cells(filaResultado, 7).Value - hora
                            wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value - hora
                        Else
                            wsResultado.Cells(filaResultado, 7).Value = wsResultado.Cells(filaResultado, 7).Value + hora
                            wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value + hora
                        End If
                    End If
                Else
                    If Cells(i, 18).Value <> 0 Then
                        'es 275.. ver si es critico
                        If Cells(i, 7).Value / Cells(i, 18).Value = 250 Then
                            If Cells(i, 6) = 2 Then
                                wsResultado.Cells(filaResultado, 6).Value = wsResultado.Cells(filaResultado, 6).Value - Cells(i, 18).Value
                                wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value - Cells(i, 18).Value
                            Else
                                wsResultado.Cells(filaResultado, 6).Value = wsResultado.Cells(filaResultado, 6).Value + Cells(i, 18).Value
                                wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value + Cells(i, 18).Value
                            End If
                        Else
                            If Cells(i, 6) = 2 Then
                                wsResultado.Cells(filaResultado, 8).Value = wsResultado.Cells(filaResultado, 8).Value - Cells(i, 18).Value
                                wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value - Cells(i, 18).Value
                            Else
                                wsResultado.Cells(filaResultado, 8).Value = wsResultado.Cells(filaResultado, 8).Value + Cells(i, 18).Value
                                wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value + Cells(i, 18).Value
                            End If
                        End If
                    Else
                        'suma todo a No critico
                        If anio > 2017 And mes > 3 Then
                            hora = Cells(i, 7).Value \ 100
                        Else
                            hora = Cells(i, 7).Value \ 150
                        End If
                        If Cells(i, 6) = 2 Then
                            wsResultado.Cells(filaResultado, 8).Value = wsResultado.Cells(filaResultado, 8).Value - hora
                            wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value - hora
                        Else
                            wsResultado.Cells(filaResultado, 8).Value = wsResultado.Cells(filaResultado, 8).Value + hora
                            wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value + hora
                        End If
                    End If
                End If
            End If
        Else
            dni = Cells(i, 12).Value
            ultMes = mes
            filaResultado = filaResultado + 1
            wsResultado.Cells(filaResultado, 1).Value = Cells(i, 12).Value
            wsResultado.Cells(filaResultado, 2).Value = Cells(i, 14).Value
            wsResultado.Cells(filaResultado, 3).Value = mes
            wsResultado.Cells(filaResultado, 4).Value = anio
            wsResultado.Cells(filaResultado, 5).Value = 0
            wsResultado.Cells(filaResultado, 6).Value = 0
            wsResultado.Cells(filaResultado, 7).Value = 0
            wsResultado.Cells(filaResultado, 8).Value = 0
            wsResultado.Cells(filaResultado, 9).Value = 0
            wsResultado.Cells(filaResultado, 10).Value = 0
            i = i - 1
        End If
    Next i
    
    
    
    For i = 2 To filaResultado
        If wsResultado.Cells(i, 10).Value > 300 Then
            'FALTARÍA controlar bien por el CUOF para 500 hs
            valorDoc = wsContenido.Cells(i, 1).Value
            'Busca en el otro archivo
            rangoTemp = "B2:B" & nFilas
            Set resultado = Range(rangoTemp).Find(What:=valorDoc, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False, _
                        SearchFormat:=False)
            'Si el resultado de la búsqueda no es vacío
            If Not resultado Is Nothing Then
                If wsResultado.Cells(i, 10).Value > 500 Then
                    wsResultado.Cells(i, 11).Value = "Supera HS. Agente 500 hs"
                    wsResultado.Cells(i, 12).Value = wsResultado.Cells(i, 10).Value - 500
                Else
                    wsResultado.Cells(i, 11).Value = "Agente 500 hs"
                End If
            Else
                wsResultado.Cells(i, 11).Value = "Supera HS"
                wsResultado.Cells(i, 12).Value = wsResultado.Cells(i, 10).Value - 300
            End If
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


Sub Control_Cuof()
    Dim i As Long
    Dim rango As Range
    Dim rangoCont As Range
    Dim resultado As Range
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim nFilasCont As Double
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet, _
        wsResultado As Excel.Worksheet
    Dim temp As String
    Dim bandera As Boolean
    Dim filaResultado As Long
    Dim fecha As Date
    Dim cuof As Integer
    Dim hsEPs As Long
    Dim hsEAc As Long
    Dim hsEAcCr As Long
    
    
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
    ActiveSheet.Name = "C_Cuof"
    Application.DisplayAlerts = True
    Set wsResultado = Worksheets("C_Cuof")
    
    Set wsContenido = wbContenido.Worksheets("Hoja1")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count
    
    MsgBox "Debe estar ordenado por Couf + Anexo + VTO", , "Atención!!"
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    filaResultado = 1
    wsResultado.Cells(filaResultado, 1).Value = "CUOF"
    wsResultado.Cells(filaResultado, 2).Value = "Anexo"
    wsResultado.Cells(filaResultado, 3).Value = "Mes"
    wsResultado.Cells(filaResultado, 4).Value = "Año"
    wsResultado.Cells(filaResultado, 5).Value = "Hs_276_Crit"
    wsResultado.Cells(filaResultado, 6).Value = "Hs_275_Crit"
    wsResultado.Cells(filaResultado, 7).Value = "Hs_276"
    wsResultado.Cells(filaResultado, 8).Value = "Hs_275"
    wsResultado.Cells(filaResultado, 9).Value = "Hs_277"
    wsResultado.Cells(filaResultado, 10).Value = "Max Hs Act.Criticas"
    wsResultado.Cells(filaResultado, 11).Value = "Max Hs Activas"
    wsResultado.Cells(filaResultado, 12).Value = "Max Hs Pasivas"
    wsResultado.Cells(filaResultado, 13).Value = "Observación"
    wsResultado.Cells(filaResultado, 14).Value = "Hs exced Act.Crit"
    wsResultado.Cells(filaResultado, 15).Value = "Hs exced Act"
    wsResultado.Cells(filaResultado, 16).Value = "Hs exced Pas"
    
    cuof = Cells(2, 20).Value
    ultAnexo = 7
    ultMes = 0
    
    For i = 2 To nFilas
        fecha = Cells(i, 16).Value
        anio = Year(fecha)
        mes = Month(fecha)
    
        If cuof = Cells(i, 20).Value And ultAnexo = Cells(i, 21).Value And ultMes = mes Then
            If Cells(i, 4) = "277" Then
                If Cells(i, 18) <> 0 Then
                    If Cells(i, 6) = 2 Then
                        wsResultado.Cells(filaResultado, 9).Value = wsResultado.Cells(filaResultado, 9).Value - Cells(i, 18).Value
                        'wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value - Cells(i, 18).Value
                    Else
                        wsResultado.Cells(filaResultado, 9).Value = wsResultado.Cells(filaResultado, 9).Value + Cells(i, 18).Value
                        'wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value + Cells(i, 18).Value
                    End If
                Else
                    If anio > 2017 And mes > 3 Then
                        hora = Cells(i, 7).Value \ 40
                    Else
                        hora = Cells(i, 7).Value \ 60
                    End If
                    If Cells(i, 6) = 2 Then
                        wsResultado.Cells(filaResultado, 9).Value = wsResultado.Cells(filaResultado, 9).Value - hora
                        'wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value - hora
                    Else
                        wsResultado.Cells(filaResultado, 9).Value = wsResultado.Cells(filaResultado, 9).Value + hora
                        'wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value + hora
                    End If
                End If
            Else
                If Cells(i, 4) = "276" Then
                    If Cells(i, 18).Value <> 0 Then
                        'ver si es critico
                        If Cells(i, 7).Value / Cells(i, 18).Value = 325 Then
                            If Cells(i, 6) = 2 Then
                                wsResultado.Cells(filaResultado, 5).Value = wsResultado.Cells(filaResultado, 5).Value - Cells(i, 18).Value
                                'wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value - Cells(i, 18).Value
                            Else
                                wsResultado.Cells(filaResultado, 5).Value = wsResultado.Cells(filaResultado, 5).Value + Cells(i, 18).Value
                                'wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value + Cells(i, 18).Value
                            End If
                        Else
                            If Cells(i, 6) = 2 Then
                                wsResultado.Cells(filaResultado, 7).Value = wsResultado.Cells(filaResultado, 7).Value - Cells(i, 18).Value
                                'wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value - Cells(i, 18).Value
                            Else
                                wsResultado.Cells(filaResultado, 7).Value = wsResultado.Cells(filaResultado, 7).Value + Cells(i, 18).Value
                                'wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value + Cells(i, 18).Value
                            End If
                        End If
                    Else
                        'suma todo a No critico
                        If anio > 2017 And mes > 3 Then
                            hora = Cells(i, 7).Value \ 150
                        Else
                            hora = Cells(i, 7).Value \ 225
                        End If
                        If Cells(i, 6) = 2 Then
                            wsResultado.Cells(filaResultado, 7).Value = wsResultado.Cells(filaResultado, 7).Value - hora
                            'wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value - hora
                        Else
                            wsResultado.Cells(filaResultado, 7).Value = wsResultado.Cells(filaResultado, 7).Value + hora
                            'wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value + hora
                        End If
                    End If
                Else
                    If Cells(i, 18).Value <> 0 Then
                        'es 275.. ver si es critico
                        If Cells(i, 7).Value / Cells(i, 18).Value = 250 Then
                            If Cells(i, 6) = 2 Then
                                wsResultado.Cells(filaResultado, 6).Value = wsResultado.Cells(filaResultado, 6).Value - Cells(i, 18).Value
                                'wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value - Cells(i, 18).Value
                            Else
                                wsResultado.Cells(filaResultado, 6).Value = wsResultado.Cells(filaResultado, 6).Value + Cells(i, 18).Value
                                'wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value + Cells(i, 18).Value
                            End If
                        Else
                            If Cells(i, 6) = 2 Then
                                wsResultado.Cells(filaResultado, 8).Value = wsResultado.Cells(filaResultado, 8).Value - Cells(i, 18).Value
                                'wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value - Cells(i, 18).Value
                            Else
                                wsResultado.Cells(filaResultado, 8).Value = wsResultado.Cells(filaResultado, 8).Value + Cells(i, 18).Value
                                'wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value + Cells(i, 18).Value
                            End If
                        End If
                    Else
                        'suma todo a No critico
                        If anio > 2017 And mes > 3 Then
                            hora = Cells(i, 7).Value \ 100
                        Else
                            hora = Cells(i, 7).Value \ 150
                        End If
                        If Cells(i, 6) = 2 Then
                            wsResultado.Cells(filaResultado, 8).Value = wsResultado.Cells(filaResultado, 8).Value - hora
                            'wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value - hora
                        Else
                            wsResultado.Cells(filaResultado, 8).Value = wsResultado.Cells(filaResultado, 8).Value + hora
                            'wsResultado.Cells(filaResultado, 10).Value = wsResultado.Cells(filaResultado, 10).Value + hora
                        End If
                    End If
                End If
            End If
        Else
            cuof = Cells(i, 20).Value
            ultAnexo = Cells(i, 21).Value
            ultMes = mes
            filaResultado = filaResultado + 1
            wsResultado.Cells(filaResultado, 1).Value = Cells(i, 20).Value
            wsResultado.Cells(filaResultado, 2).Value = Cells(i, 21).Value
            wsResultado.Cells(filaResultado, 3).Value = mes
            wsResultado.Cells(filaResultado, 4).Value = anio
            wsResultado.Cells(filaResultado, 5).Value = 0
            wsResultado.Cells(filaResultado, 6).Value = 0
            wsResultado.Cells(filaResultado, 7).Value = 0
            wsResultado.Cells(filaResultado, 8).Value = 0
            wsResultado.Cells(filaResultado, 9).Value = 0
            i = i - 1
        End If
    Next i
    
    For i = 2 To filaResultado
        bandera = True
        For j = 3 To 115
            If wsContenido.Cells(j, 2).Value = wsResultado.Cells(i, 1).Value Then
                If wsContenido.Cells(j, 3).Value = wsResultado.Cells(i, 2).Value Then
                    hsEAc = 0
                    hsEAcCr = 0
                    hsEPs = 0
                    
                    If wsResultado.Cells(i, 4).Value > 2017 And wsResultado.Cells(i, 3).Value > 3 Then
                        wsResultado.Cells(i, 10).Value = wsContenido.Cells(j, 8).Value
                        wsResultado.Cells(i, 11).Value = wsContenido.Cells(j, 9).Value
                        wsResultado.Cells(i, 12).Value = wsContenido.Cells(j, 10).Value
                        hsEAcCr = wsResultado.Cells(i, 5).Value + wsResultado.Cells(i, 6).Value - wsResultado.Cells(i, 10).Value
                    Else
                        wsResultado.Cells(i, 11).Value = wsContenido.Cells(j, 6).Value
                        wsResultado.Cells(i, 12).Value = wsContenido.Cells(j, 7).Value
                        hsEAcCr = wsResultado.Cells(i, 5).Value + wsResultado.Cells(i, 6).Value
                    End If
                    hsEAc = wsResultado.Cells(i, 7).Value + wsResultado.Cells(i, 8).Value - wsResultado.Cells(i, 11).Value
                    hsEPs = wsResultado.Cells(i, 9).Value - wsResultado.Cells(i, 12).Value
                    
                    If hsEAcCr > 0 Then
                        wsResultado.Cells(i, 13).Value = "Controlar"
                        wsResultado.Cells(i, 14).Value = hsEAcCr
                    End If
                    If hsEAc > 0 Then
                        wsResultado.Cells(i, 13).Value = "Controlar"
                        wsResultado.Cells(i, 15).Value = hsEAc
                    End If
                    If hsEPs > 0 Then
                        wsResultado.Cells(i, 13).Value = "Controlar"
                        wsResultado.Cells(i, 16).Value = hsEPs
                    End If
                    
                    bandera = False
                    j = 115
                End If
            End If
        Next j
        If bandera Then
            wsResultado.Cells(i, 13).Value = "No encontró Cuof+Anexo"
            wsResultado.Cells(i, 14).Value = wsResultado.Cells(i, 5).Value + wsResultado.Cells(i, 6).Value
            wsResultado.Cells(i, 15).Value = wsResultado.Cells(i, 7).Value + wsResultado.Cells(i, 8).Value
            wsResultado.Cells(i, 16).Value = wsResultado.Cells(i, 9).Value
        End If
    Next i
    
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub



Sub Control_Cuof_2()
    Dim i As Long
    Dim rango As Range
    Dim rangoCont As Range
    Dim resultado As Range
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim nFilasCont As Double
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet, _
        wbContCuof As Workbook, _
        wsContCuof As Excel.Worksheet, _
        wsResultado As Excel.Worksheet
    Dim temp As String
    Dim bandera As Boolean
    Dim filaResultado As Long
    Dim fecha As Date
    Dim cuof As Integer
    Dim hsEPs As Long
    Dim hsEAc As Long
    Dim hsEAcCr As Long
    
    contenido = InputBox("Ingrese el nombre del archivo Carga:", "Abrir", "Archivo.xlsx")
    If contenido <> "" Then
        On Error GoTo ControlErrorOpen
        Set wbContCuof = Workbooks.Open(ActiveWorkbook.Path & "\" & contenido)
    Else
        Exit Sub
    End If
    
    'Activar este libro
    ThisWorkbook.Activate
       
    'Agrego la nueva hoja
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "C_Cuof"
    Application.DisplayAlerts = True
    Set wsResultado = Worksheets("C_Cuof")
    
    Set wsContCuof = wbContCuof.Worksheets("Hoja1")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count
    
    MsgBox "Deben estar ordenados por DNI + VTO", , "Atención!!"
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContCuof.UsedRange
    nFilasCont = rangoCont.Rows.Count
    nColumnasCont = rangoCont.Columns.Count
    
    filaResultado = 1
    wsResultado.Cells(filaResultado, 1).Value = "CUOF"
    wsResultado.Cells(filaResultado, 2).Value = "Anexo"
    wsResultado.Cells(filaResultado, 3).Value = "Mes"
    wsResultado.Cells(filaResultado, 4).Value = "Año"
    wsResultado.Cells(filaResultado, 5).Value = "Hs_276_Crit"
    wsResultado.Cells(filaResultado, 6).Value = "Hs_275_Crit"
    wsResultado.Cells(filaResultado, 7).Value = "Hs_276"
    wsResultado.Cells(filaResultado, 8).Value = "Hs_275"
    wsResultado.Cells(filaResultado, 9).Value = "Hs_277"
    wsResultado.Cells(filaResultado, 10).Value = "Max Hs Act.Criticas"
    wsResultado.Cells(filaResultado, 11).Value = "Max Hs Activas"
    wsResultado.Cells(filaResultado, 12).Value = "Max Hs Pasivas"
    wsResultado.Cells(filaResultado, 13).Value = "Observación"
    wsResultado.Cells(filaResultado, 14).Value = "Hs exced Act.Crit"
    wsResultado.Cells(filaResultado, 15).Value = "Hs exced Act"
    wsResultado.Cells(filaResultado, 16).Value = "Hs exced Pas"
    
    valorDoc = "0"
    
    For i = 2 To nFilasCont
        If Cells(i, 6).Value = 0 Then
            fecha = Cells(i, 16).Value
            anio = Year(fecha)
            mes = Month(fecha)
            
            If valorDoc <> Cells(i, 12).Value Then
                valorDoc = Cells(i, 12).Value
                'Busca en el otro archivo
                rangoTemp = "H2:H" & nFilas
                Set resultado = wsContCuof.Range(rangoTemp).Find(What:=valorDoc, _
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
                    
                    Do While wsContCuof.Cells(j, 8).Value = wsContCuof.Cells(j - 1, 8).Value
                        j = j - 1
                    Loop
                    
                    tempj = j
                    
                    Cells(i, nColumnas + 1).Value = "No encontrado"
                    cuof = Cells(i, 20).Value
                    anexo = Cells(i, 21).Value
                    Do
                        If wsContCuof.Cells(j, 11).Value = mes And wsContCuof.Cells(j, 12).Value = anio Then
                            If wsContCuof.Cells(j, 15).Value = Cells(i, 4).Value And wsContCuof.Cells(j, 17).Value = Cells(i, 7).Value Then
                                cuof = wsContCuof.Cells(j, 4).Value
                                Cells(i, nColumnas + 1).Value = ""
                                anexo = wsContCuof.Cells(j, 5).Value
                                'wsContCuof.Cells(j, nColumnasCont + 1).Value = "Procesado"
                                j = j + 10
                            End If
                        End If
                        j = j + 1
                    Loop While valorDoc = wsContCuof.Cells(j, 8).Value
                Else
                    Cells(i, nColumnas + 1).Value = "No existe DNI en Archivo de Carga"
                    cuof = Cells(i, 20).Value
                    anexo = Cells(i, 21).Value
                End If
            Else
                j = tempj
                Cells(i, nColumnas + 1).Value = "No encontrado"
                cuof = Cells(i, 20).Value
                anexo = Cells(i, 21).Value
                Do
                    If wsContCuof.Cells(j, 11).Value = mes And wsContCuof.Cells(j, 12).Value = anio Then
                        If wsContCuof.Cells(j, 15).Value = Cells(i, 4).Value And wsContCuof.Cells(j, 17).Value = Cells(i, 7).Value Then
                            cuof = wsContCuof.Cells(j, 4).Value
                            Cells(i, nColumnas + 1).Value = ""
                            j = j + 10
                            anexo = wsContCuof.Cells(j, 5).Value
                        End If
                    End If
                    j = j + 1
                Loop While valorDoc = wsContCuof.Cells(j, 8).Value
            End If
            
            'Cargar hora
            k = 2
            bandera = True
            Do While k <= filaResultado
                If wsResultado.Cells(k, 1).Value = cuof And wsResultado.Cells(k, 2).Value = anexo Then
                    If wsResultado.Cells(k, 3).Value = mes And wsResultado.Cells(k, 4).Value = anio Then
                        p = k
                        k = filaResultado
                        bandera = False
                    End If
                End If
                k = k + 1
            Loop
            If bandera Then
                filaResultado = filaResultado + 1
                wsResultado.Cells(filaResultado, 1).Value = cuof
                wsResultado.Cells(filaResultado, 2).Value = anexo
                wsResultado.Cells(filaResultado, 3).Value = mes
                wsResultado.Cells(filaResultado, 4).Value = anio
                wsResultado.Cells(filaResultado, 5).Value = 0
                wsResultado.Cells(filaResultado, 6).Value = 0
                wsResultado.Cells(filaResultado, 7).Value = 0
                wsResultado.Cells(filaResultado, 8).Value = 0
                wsResultado.Cells(filaResultado, 9).Value = 0
                p = filaResultado
            End If
            
            If Cells(i, 4) = "277" Then
                If Cells(i, 18) <> 0 Then
                    If Cells(i, 6) = 2 Then
                        wsResultado.Cells(p, 9).Value = wsResultado.Cells(p, 9).Value - Cells(i, 18).Value
                    Else
                        wsResultado.Cells(p, 9).Value = wsResultado.Cells(p, 9).Value + Cells(i, 18).Value
                    End If
                Else
                    If anio > 2017 And mes > 3 Then
                        hora = Cells(i, 7).Value \ 40
                    Else
                        hora = Cells(i, 7).Value \ 60
                    End If
                    If Cells(i, 6) = 2 Then
                        wsResultado.Cells(p, 9).Value = wsResultado.Cells(p, 9).Value - hora
                    Else
                        wsResultado.Cells(p, 9).Value = wsResultado.Cells(p, 9).Value + hora
                    End If
                End If
            Else
                If Cells(i, 4) = "276" Then
                    If Cells(i, 18).Value <> 0 Then
                        'ver si es critico
                        If Cells(i, 7).Value / Cells(i, 18).Value = 325 Then
                            If Cells(i, 6) = 2 Then
                                wsResultado.Cells(p, 5).Value = wsResultado.Cells(p, 5).Value - Cells(i, 18).Value
                            Else
                                wsResultado.Cells(p, 5).Value = wsResultado.Cells(p, 5).Value + Cells(i, 18).Value
                            End If
                        Else
                            If Cells(i, 6) = 2 Then
                                wsResultado.Cells(p, 7).Value = wsResultado.Cells(p, 7).Value - Cells(i, 18).Value
                            Else
                                wsResultado.Cells(p, 7).Value = wsResultado.Cells(p, 7).Value + Cells(i, 18).Value
                            End If
                        End If
                    Else
                        'suma todo a No critico
                        If anio > 2017 And mes > 3 Then
                            hora = Cells(i, 7).Value \ 150
                        Else
                            hora = Cells(i, 7).Value \ 225
                        End If
                        If Cells(i, 6) = 2 Then
                            wsResultado.Cells(p, 7).Value = wsResultado.Cells(p, 7).Value - hora
                        Else
                            wsResultado.Cells(p, 7).Value = wsResultado.Cells(p, 7).Value + hora
                        End If
                    End If
                Else
                    If Cells(i, 18).Value <> 0 Then
                        'es 275.. ver si es critico
                        If Cells(i, 7).Value / Cells(i, 18).Value = 250 Then
                            If Cells(i, 6) = 2 Then
                                wsResultado.Cells(p, 6).Value = wsResultado.Cells(p, 6).Value - Cells(i, 18).Value
                            Else
                                wsResultado.Cells(p, 6).Value = wsResultado.Cells(p, 6).Value + Cells(i, 18).Value
                            End If
                        Else
                            If Cells(i, 6) = 2 Then
                                wsResultado.Cells(p, 8).Value = wsResultado.Cells(p, 8).Value - Cells(i, 18).Value
                            Else
                                wsResultado.Cells(p, 8).Value = wsResultado.Cells(p, 8).Value + Cells(i, 18).Value
                            End If
                        End If
                    Else
                        'suma todo a No critico
                        If anio > 2017 And mes > 3 Then
                            hora = Cells(i, 7).Value \ 100
                        Else
                            hora = Cells(i, 7).Value \ 150
                        End If
                        If Cells(i, 6) = 2 Then
                            wsResultado.Cells(p, 8).Value = wsResultado.Cells(p, 8).Value - hora
                        Else
                            wsResultado.Cells(p, 8).Value = wsResultado.Cells(p, 8).Value + hora
                        End If
                    End If
                End If
            End If
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub



Sub Control_Cuof_2_SegParte()
    Dim i As Long
    Dim rango As Range
    Dim rangoCont As Range
    Dim resultado As Range
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim nFilasCont As Double
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet, _
        wbContCuof As Workbook, _
        wsContCuof As Excel.Worksheet, _
        wsResultado As Excel.Worksheet
    Dim temp As String
    Dim bandera As Boolean
    Dim filaResultado As Long
    Dim fecha As Date
    Dim cuof As Integer
    Dim hsEPs As Long
    Dim hsEAc As Long
    Dim hsEAcCr As Long
    
    
    'Indicar el libro de excel CONTENIDO y control de errores
    contenido = InputBox("Ingrese el nombre del archivo Oficinas:", "Abrir", "Archivo.xlsx")
    If contenido <> "" Then
        On Error GoTo ControlErrorOpen
        Set wbContenido = Workbooks.Open(ActiveWorkbook.Path & "\" & contenido)
    Else
        Exit Sub
    End If
    
    'Activar este libro
    ThisWorkbook.Activate
    Set wsResultado = Worksheets("C_Cuof")
    Set wsContenido = wbContenido.Worksheets("Hoja1")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsResultado.UsedRange
    filaResultado = rangoCont.Rows.Count
    
    For i = 2 To filaResultado
        bandera = True
        For j = 3 To 115
        sda = wsContenido.Cells(j, 2).Value
        sdawdaw = wsResultado.Cells(i, 1).Value
            If wsContenido.Cells(j, 2).Value = wsResultado.Cells(i, 1).Value Then
                If wsContenido.Cells(j, 3).Value = wsResultado.Cells(i, 2).Value Then
                    hsEAc = 0
                    hsEAcCr = 0
                    hsEPs = 0
                    
                    If wsResultado.Cells(i, 4).Value > 2017 And wsResultado.Cells(i, 3).Value > 3 Then
                        wsResultado.Cells(i, 10).Value = wsContenido.Cells(j, 8).Value
                        wsResultado.Cells(i, 11).Value = wsContenido.Cells(j, 9).Value
                        wsResultado.Cells(i, 12).Value = wsContenido.Cells(j, 10).Value
                        hsEAcCr = wsResultado.Cells(i, 5).Value + wsResultado.Cells(i, 6).Value - wsResultado.Cells(i, 10).Value
                    Else
                        wsResultado.Cells(i, 11).Value = wsContenido.Cells(j, 6).Value
                        wsResultado.Cells(i, 12).Value = wsContenido.Cells(j, 7).Value
                        hsEAcCr = wsResultado.Cells(i, 5).Value + wsResultado.Cells(i, 6).Value
                    End If
                    hsEAc = wsResultado.Cells(i, 7).Value + wsResultado.Cells(i, 8).Value - wsResultado.Cells(i, 11).Value
                    hsEPs = wsResultado.Cells(i, 9).Value - wsResultado.Cells(i, 12).Value
                    
                    If hsEAcCr > 0 Then
                        wsResultado.Cells(i, 13).Value = "Controlar"
                        wsResultado.Cells(i, 14).Value = hsEAcCr
                    End If
                    If hsEAc > 0 Then
                        wsResultado.Cells(i, 13).Value = "Controlar"
                        wsResultado.Cells(i, 15).Value = hsEAc
                    End If
                    If hsEPs > 0 Then
                        wsResultado.Cells(i, 13).Value = "Controlar"
                        wsResultado.Cells(i, 16).Value = hsEPs
                    End If
                    
                    bandera = False
                    j = 115
                End If
            End If
        Next j
        If bandera Then
            wsResultado.Cells(i, 13).Value = "No encontró Cuof+Anexo"
            wsResultado.Cells(i, 14).Value = wsResultado.Cells(i, 5).Value + wsResultado.Cells(i, 6).Value
            wsResultado.Cells(i, 15).Value = wsResultado.Cells(i, 7).Value + wsResultado.Cells(i, 8).Value
            wsResultado.Cells(i, 16).Value = wsResultado.Cells(i, 9).Value
        End If
    Next i
    
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub



