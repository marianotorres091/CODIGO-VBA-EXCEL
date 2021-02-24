Attribute VB_Name = "Módulo2"
Sub Control_Anual()
    Dim pagos(1 To 12) As Integer
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet
    Dim valorDoc As String
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim nColumnasCont As Long
    Dim nFilasCont As Long
    Dim i As Long
    Dim mes As Integer
    Dim mesesPagos As Integer
    Dim importeTotal As Double
    
    
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

    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    nColumnasCont = rangoCont.Columns.Count
    
    For i = 1 To 12
       pagos(i) = 0
    Next i
    valorDoc = Cells(2, 12).Value
    importeTotal = 0
    mesesPagos = 0
    
    For i = 2 To nFilas
        If valorDoc = Cells(i, 12).Value Then
            'Trato la persona
            mes = 0
            If Cells(i, 6).Value = 0 Then
                mes = Cells(i, 2).Value
            Else
                valorVto = Cells(i, 16).Value
                For m = 2 To Len(valorVto)
                    If Mid(valorVto, m, 1) = "/" Then
                        mes = Mid(valorVto, m + 1, 1)
                        m = m + 2
                        If Mid(valorVto, m, 1) <> "/" Then
                            mes = mes & Mid(valorVto, m, 1)
                        End If
                        m = 10
                    End If
                Next m
            End If
            pagos(mes) = pagos(mes) + 1
            mesesPagos = mesesPagos + 1
            importeTotal = importeTotal + Cells(i, 7).Value
        Else
            'Copio en archivo
            i = i - 1
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
                
                wsContenido.Cells(j, nColumnasCont + 2).Value = importeTotal
                wsContenido.Cells(j, nColumnasCont + 1).Value = mesesPagos
            Else
                nFilasCont = nFilasCont + 1
                wsContenido.Cells(nFilasCont, 1).Value = Cells(i, 12).Value
                wsContenido.Cells(nFilasCont, 2).Value = Cells(i, 14).Value
                wsContenido.Cells(nFilasCont, nColumnasCont + 2).Value = importeTotal
                wsContenido.Cells(nFilasCont, nColumnasCont + 1).Value = mesesPagos
            End If
            'Reinicio contadores
            For j = 1 To 12
               pagos(j) = 0
            Next j
            valorDoc = Cells(i + 1, 12).Value
            importeTotal = 0
            mesesPagos = 0
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


Sub Control_Anual_Mes()
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet
    Dim valorDoc As String
    Dim nFilas As Long
    Dim nFilasCont As Long
    Dim i As Long
    Dim mes As Integer
    
    
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

    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    wsContenido.Cells(1, 1).Value = "JUR"
    wsContenido.Cells(1, 2).Value = "DNI"
    wsContenido.Cells(1, 3).Value = "Nombre"
    wsContenido.Cells(1, 4).Value = "Enero"
    wsContenido.Cells(1, 5).Value = "Febrero"
    wsContenido.Cells(1, 6).Value = "Marzo"
    wsContenido.Cells(1, 7).Value = "Abril"
    wsContenido.Cells(1, 8).Value = "Mayo"
    wsContenido.Cells(1, 9).Value = "Junio"
    wsContenido.Cells(1, 10).Value = "Julio"
    wsContenido.Cells(1, 11).Value = "Agosto"
    wsContenido.Cells(1, 12).Value = "Septiembre"
    wsContenido.Cells(1, 13).Value = "Octubre"
    wsContenido.Cells(1, 14).Value = "Noviembre"
    wsContenido.Cells(1, 15).Value = "Diciembre"
    nFilasCont = 2
    
    valorDoc = Cells(2, 12).Value
    wsContenido.Cells(nFilasCont, 1).Value = Cells(2, 8).Value
    wsContenido.Cells(nFilasCont, 2).Value = Cells(2, 12).Value
    wsContenido.Cells(nFilasCont, 3).Value = Cells(2, 14).Value
    
    For i = 2 To nFilas
        If valorDoc = Cells(i, 12).Value Then
            'Trato la persona
            mes = 0
            anio = 0
            
            valorVto = Cells(i, 16).Value
            For m = 2 To Len(valorVto)
                If Mid(valorVto, m, 1) = "/" Then
                    mes = Mid(valorVto, m + 1, 1)
                    m = m + 2
                    If Mid(valorVto, m, 1) <> "/" Then
                        mes = mes & Mid(valorVto, m, 1)
                        m = m + 1
                    End If
                    anio = Mid(valorVto, m + 1, 4)
                    m = 10
                End If
            Next m
            
            If anio = 2017 Then
                If mes < Cells(i, 2).Value Then
                    If Cells(i, 6).Value = 2 Then
                        wsContenido.Cells(nFilasCont, mes + 3).Value = wsContenido.Cells(nFilasCont, mes + 3).Value - Cells(i, 7).Value
                    Else
                        wsContenido.Cells(nFilasCont, mes + 3).Value = wsContenido.Cells(nFilasCont, mes + 3).Value + Cells(i, 7).Value
                    End If
                Else
                    If Cells(i, 6).Value = 2 Then
                        wsContenido.Cells(nFilasCont, Cells(i, 2).Value + 3).Value = wsContenido.Cells(nFilasCont, Cells(i, 2).Value + 3).Value - Cells(i, 7).Value
                    Else
                        wsContenido.Cells(nFilasCont, Cells(i, 2).Value + 3).Value = wsContenido.Cells(nFilasCont, Cells(i, 2).Value + 3).Value + Cells(i, 7).Value
                    End If
                End If
            Else
                If anio > 2017 Then
                    If Cells(i, 6).Value = 2 Then
                        wsContenido.Cells(nFilasCont, Cells(i, 2).Value + 3).Value = wsContenido.Cells(nFilasCont, Cells(i, 2).Value + 3).Value - Cells(i, 7).Value
                    Else
                        wsContenido.Cells(nFilasCont, Cells(i, 2).Value + 3).Value = wsContenido.Cells(nFilasCont, Cells(i, 2).Value + 3).Value + Cells(i, 7).Value
                    End If
                End If
            End If
        Else
            nFilasCont = nFilasCont + 1
            wsContenido.Cells(nFilasCont, 1).Value = Cells(i, 8).Value
            wsContenido.Cells(nFilasCont, 2).Value = Cells(i, 12).Value
            wsContenido.Cells(nFilasCont, 3).Value = Cells(i, 14).Value
            'Reinicio contadores
            valorDoc = Cells(i, 12).Value
            'Trato la persona
            mes = 0
            anio = 0
            
            valorVto = Cells(i, 16).Value
            For m = 2 To Len(valorVto)
                If Mid(valorVto, m, 1) = "/" Then
                    mes = Mid(valorVto, m + 1, 1)
                    m = m + 2
                    If Mid(valorVto, m, 1) <> "/" Then
                        mes = mes & Mid(valorVto, m, 1)
                        m = m + 1
                    End If
                    anio = Mid(valorVto, m + 1, 4)
                    m = 10
                End If
            Next m

            If anio = 2017 Then
                If mes < Cells(i, 2).Value Then
                    If Cells(i, 6).Value = 2 Then
                        wsContenido.Cells(nFilasCont, mes + 3).Value = wsContenido.Cells(nFilasCont, mes + 3).Value - Cells(i, 7).Value
                    Else
                        wsContenido.Cells(nFilasCont, mes + 3).Value = wsContenido.Cells(nFilasCont, mes + 3).Value + Cells(i, 7).Value
                    End If
                Else
                    If Cells(i, 6).Value = 2 Then
                        wsContenido.Cells(nFilasCont, Cells(i, 2).Value + 3).Value = wsContenido.Cells(nFilasCont, Cells(i, 2).Value + 3).Value - Cells(i, 7).Value
                    Else
                        wsContenido.Cells(nFilasCont, Cells(i, 2).Value + 3).Value = wsContenido.Cells(nFilasCont, Cells(i, 2).Value + 3).Value + Cells(i, 7).Value
                    End If
                End If
            Else
                If anio > 2017 Then
                    If Cells(i, 6).Value = 2 Then
                        wsContenido.Cells(nFilasCont, Cells(i, 2).Value + 3).Value = wsContenido.Cells(nFilasCont, Cells(i, 2).Value + 3).Value - Cells(i, 7).Value
                    Else
                        wsContenido.Cells(nFilasCont, Cells(i, 2).Value + 3).Value = wsContenido.Cells(nFilasCont, Cells(i, 2).Value + 3).Value + Cells(i, 7).Value
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


Sub Control_Diferencia()
    Dim valorDoc As String
    Dim nFilas As Long
    Dim nColumnas As Integer
    Dim i As Long
    Dim bandera As Boolean
    Dim monto As Double
    
    'Activar este libro
    ThisWorkbook.Activate

    'Regresa el control a la hoja de origen
    Sheets(1).Select

    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    Cells(1, nColumnas + 1).Value = "Observación"
    Cells(1, nColumnas + 2).Value = "Diferencia Mayor"
            
    For i = 2 To nFilas
        bandera = False
        monto = 0
        For j = 4 To nColumnas
            If Cells(i, j - 1).Value > Cells(i, j).Value And Cells(i, j).Value > 0 Then
                If (Cells(i, j - 1).Value - Cells(i, j).Value) > monto And (Cells(i, j - 1).Value - Cells(i, j).Value) > 5 Then
                    monto = Cells(i, j - 1).Value - Cells(i, j).Value
                    bandera = True
                End If
            End If
        Next j
        If bandera Then
            Cells(i, nColumnas + 1).Value = "Controlar"
            Cells(i, nColumnas + 2).Value = monto
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub



Sub Control_Anual_Mes_Completo()
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet
    Dim valorDoc As String
    Dim nFilas As Long
    Dim nFilasCont As Long
    Dim i As Long
    Dim mes As Integer
    Dim anio As Integer
    
    
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

    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    wsContenido.Cells(1, 1).Value = "DNI"
    wsContenido.Cells(1, 2).Value = "Nombre"
    wsContenido.Cells(1, 3).Value = "Enero-15"
    wsContenido.Cells(1, 4).Value = "Febrero-15"
    wsContenido.Cells(1, 5).Value = "Marzo-15"
    wsContenido.Cells(1, 6).Value = "Abril-15"
    wsContenido.Cells(1, 7).Value = "Mayo-15"
    wsContenido.Cells(1, 8).Value = "Junio-15"
    wsContenido.Cells(1, 9).Value = "Julio-15"
    wsContenido.Cells(1, 10).Value = "Agosto-15"
    wsContenido.Cells(1, 11).Value = "Septiembre-15"
    wsContenido.Cells(1, 12).Value = "Octubre-15"
    wsContenido.Cells(1, 13).Value = "Noviembre-15"
    wsContenido.Cells(1, 14).Value = "Diciembre-15"
    wsContenido.Cells(1, 15).Value = "Enero-16"
    wsContenido.Cells(1, 16).Value = "Febrero-16"
    wsContenido.Cells(1, 17).Value = "Marzo-16"
    wsContenido.Cells(1, 18).Value = "Abril-16"
    wsContenido.Cells(1, 19).Value = "Mayo-16"
    wsContenido.Cells(1, 20).Value = "Junio-16"
    wsContenido.Cells(1, 21).Value = "Julio-16"
    wsContenido.Cells(1, 22).Value = "Agosto-16"
    wsContenido.Cells(1, 23).Value = "Septiembre-16"
    wsContenido.Cells(1, 24).Value = "Octubre-16"
    wsContenido.Cells(1, 25).Value = "Noviembre-16"
    wsContenido.Cells(1, 26).Value = "Diciembre-16"
    wsContenido.Cells(1, 27).Value = "Enero-17"
    wsContenido.Cells(1, 28).Value = "Febrero-17"
    wsContenido.Cells(1, 29).Value = "Marzo-17"
    wsContenido.Cells(1, 30).Value = "Abril-17"
    wsContenido.Cells(1, 31).Value = "Mayo-17"
    wsContenido.Cells(1, 32).Value = "Junio-17"
    wsContenido.Cells(1, 33).Value = "Julio-17"
    wsContenido.Cells(1, 34).Value = "Agosto-17"
    wsContenido.Cells(1, 35).Value = "Septiembre-17"
    wsContenido.Cells(1, 36).Value = "Octubre-17"
    wsContenido.Cells(1, 37).Value = "Noviembre-17"
    wsContenido.Cells(1, 38).Value = "Diciembre-17"
    nFilasCont = 2
    
    valorDoc = Cells(2, 12).Value
    wsContenido.Cells(nFilasCont, 1).Value = Cells(2, 12).Value
    wsContenido.Cells(nFilasCont, 2).Value = Cells(2, 14).Value
    
    For i = 2 To nFilas
        If valorDoc = Cells(i, 12).Value Then
            'Trato la persona
            mes = 0
            If Cells(i, 6).Value = 0 Then
                mes = Cells(i, 2).Value
                anio = Cells(i, 1).Value
            Else
                valorVto = Cells(i, 16).Value
                For m = 2 To Len(valorVto)
                    If Mid(valorVto, m, 1) = "/" Then
                        mes = Mid(valorVto, m + 1, 1)
                        m = m + 2
                        If Mid(valorVto, m, 1) <> "/" Then
                            mes = mes & Mid(valorVto, m, 1)
                            m = m + 1
                        End If
                        anio = Mid(valorVto, m + 1, 4)
                        m = 10
                    End If
                Next m
            End If
            If anio = 2015 Then
                anio = 2
            Else
                If anio = 2016 Then
                    anio = 14
                Else
                    anio = 26
                End If
            End If
            wsContenido.Cells(nFilasCont, mes + anio).Value = wsContenido.Cells(nFilasCont, mes + anio).Value + Cells(i, 7).Value
        Else
            nFilasCont = nFilasCont + 1
            wsContenido.Cells(nFilasCont, 1).Value = Cells(i, 12).Value
            wsContenido.Cells(nFilasCont, 2).Value = Cells(i, 14).Value
            'Reinicio contadores
            valorDoc = Cells(i + 1, 12).Value
            'Trato la persona
            mes = 0
            anio = 0
            If Cells(i, 6).Value = 0 Then
                mes = Cells(i, 2).Value
                anio = Cells(i, 1).Value
            Else
                valorVto = Cells(i, 16).Value
                For m = 2 To Len(valorVto)
                    If Mid(valorVto, m, 1) = "/" Then
                        mes = Mid(valorVto, m + 1, 1)
                        m = m + 2
                        If Mid(valorVto, m, 1) <> "/" Then
                            mes = mes & Mid(valorVto, m, 1)
                        End If
                        anio = Mid(valorVto, m + 1, 4)
                        m = 10
                    End If
                Next m
            End If
            If anio = 2015 Then
                anio = 2
            Else
                If anio = 2016 Then
                    anio = 14
                Else
                    anio = 26
                End If
            End If
            wsContenido.Cells(nFilasCont, mes + anio).Value = wsContenido.Cells(nFilasCont, mes + anio).Value + Cells(i, 7).Value
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub

Sub Correccion_Error()
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet
    Dim valorDoc As String
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim nColumnasCont As Long
    Dim nFilasCont As Long
    Dim i As Long
    
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

    Set wsContenido = wbContenido.Worksheets("JUR3")
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
    
    For i = 2 To nFilas
        valorDoc = Cells(i, 1).Value
        rangoTemp = "B2:B" & nFilasCont
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
            
            wsContenido.Cells(j, 13).Value = wsContenido.Cells(j, 13).Value - Cells(i, 3).Value
            wsContenido.Cells(j, 14).Value = wsContenido.Cells(j, 14).Value + Cells(i, 3).Value
            wsContenido.Cells(j, 15).Value = "Modificado"
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub

