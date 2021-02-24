Attribute VB_Name = "Módulo11"
Sub Insertar_Vto()
    'Copio formato de una celda y pego en la nueva
    Range("A1").Copy
    Range("T1").PasteSpecial Paste:=xlFormats
    Application.CutCopyMode = False
    Range("T1").Value = "Vto"
    
    Dim j As Long
    Dim celdaDestino As String
    Dim celdaAge As String
    Dim celdaMonth As String
    Dim valorVto As String
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim rango As Range
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count
        
    'Asignar a cada celda de la columna VTO el valor correspondiente
    For j = 2 To nFilas
        valorVto = Cells(j, 13).Value & Cells(j, 12).Value
        Cells(j, 20).Value = valorVto
    Next j
    
    'Mostrar msj para confirmar
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    
End Sub

Sub Cambiar_Vto()
    Dim i As Long
    Dim valorVto As String
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim rango As Range
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count
    
    For i = 1 To nFilas
        valorVto = Cells(i, 12).Value
        For m = 1 To Len(valorVto)
            If Mid(valorVto, m, 1) = "/" Then
                tempMes = Mid(valorVto, m + 1, 1)
                m = m + 2
                If Mid(valorVto, m, 1) <> "/" Then
                    tempMes = tempMes & Mid(valorVto, m, 1)
                    m = m + 1
                End If
                tempAnio = Mid(valorVto, m + 1, 4)
                m = 10
            End If
        Next m
        Cells(i, nColumnas + 1).Value = tempMes & tempAnio
    Next i
End Sub

Sub Insertar_SAC()
    Dim j As Long
    Dim i As Long
    Dim celdaDif As String
    Dim celdaMonth As String
    Dim valorVto As String
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim rango As Range
    Dim temporal As Variant
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count
    
    j = 2
    Do While Not IsEmpty(Cells(j, "D"))
        'Identificar y agregar filas cuando corresponda para el SAC
        If (Cells(j, 13).Value = 6 Or Cells(j, 13).Value = 12) And Cells(j, 19).Value > 0 Then
            
            j = j + 1
            Rows(j).Insert
            'Copio la fila completa en la nueva
            For i = 1 To nColumnas
                temporal = Cells((j - 1), i).Value
                Cells(j, i).Value = temporal
            Next i
            
            'Modifica los valores que se deben cambiar
            Cells(j, 15).Value = 316
            Cells(j, 16).Value = 0
            Cells(j, 19).Value = Cells(j, 19).Value / 2
        End If
        j = j + 1
    Loop
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub


Sub Corregir_SAC()
    Dim j As Long
    Dim i As Long
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim rango As Range
    Dim temporal As Double
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count
    
    doc = Cells(1, 5).Value
    mes = Cells(1, 12).Value
    temporal = 0
    For i = 2 To nFilas
        If Cells(i, 8).Value <> 316 Then
            If doc = Cells(i, 5).Value And mes = Cells(i, 12).Value Then
                If Cells(i, 9).Value = 2 Then
                    temporal = temporal - Cells(i, 11).Value
                Else
                    temporal = temporal + Cells(i, 11).Value
                End If
            Else
                doc = Cells(i, 5).Value
                mes = Cells(i, 12).Value
                temporal = 0
                i = i - 1
            End If
        Else
            If doc = Cells(i, 5).Value And mes = Cells(i, 12).Value Then
                Cells(i, 11).Value = temporal / 2
            Else
                Cells(i, nColumnas + 1).Value = "ERROR"
            End If
        End If
    Next i
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub


Sub Insertar_SAC_Jur()
    Dim j As Long
    Dim i As Long
    Dim celdaDif As String
    Dim celdaMonth As String
    Dim valorVto As String
    Dim celdaJur As String
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim rango As Range
    Dim rangoJur As String
    Dim resultado As Range
    Dim temporal As Variant
    Dim valorJur As Integer
    Dim tempJur As String
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count
    
    j = 2
    'valorJur = CInt(InputBox("Ingrese el número de la Juridiscción:", "Calcular SAC"))
    tempJur = InputBox("Ingrese el número de la Juridiscción:", "Calcular SAC")
    
    'condicional si tempJur es numero
    If IsNumeric(tempJur) Then
        valorJur = CInt(tempJur)
        rangoJur = "A1:A" & nFilas
        'Se podría avisar si la JUR no existe
        Set resultado = Range(rangoJur).Find(What:=valorJur, _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False, _
                    SearchFormat:=False)
        'Si el resultado de la búsqueda no es vacío
        If Not resultado Is Nothing Then
                primerResultado = resultado.Address
            Do
                'Se obtiene el valor de j
                celdaJur = resultado.Address
                tempJur = ""
                For i = 1 To Len(celdaJur)
                    If IsNumeric(Mid(celdaJur, i, 1)) Then
                        tempJur = tempJur & Mid(celdaJur, i, 1)
                    End If
                Next i
                j = tempJur
                'Identifica y agrega filas cuando corresponda para el SAC
                If Range(celdaJur).Value = valorJur And (Cells(j, 13).Value = 6 Or Cells(j, 13).Value = 12) And Cells(j, 19).Value > 0 Then
                    j = j + 1
                    Rows(j).Insert
                    'Copio la fila completa en la nueva
                    For i = 1 To nColumnas
                        temporal = Cells((j - 1), i).Value
                        Cells(j, i).Value = temporal
                    Next i
                    'Modifica los valores que se deben cambiar
                    Cells(j, 15).Value = 316
                    Cells(j, 16).Value = 0
                    Cells(j, 19).Value = Cells(j, 19).Value / 2
                    'Salto una linea, la que se agregó
                    Set resultado = Range(rangoJur).FindNext(resultado)
                End If
                Set resultado = Range(rangoJur).FindNext(resultado)
            Loop While Not resultado Is Nothing And primerResultado <> resultado.Address
            
            MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
        Else
            MsgBox "El número de Juridiscción ingresado no existe.", , "Atención"
        End If
    Else
        If tempJur <> "" Then
            MsgBox "Debe ingresar un número de Juridiscción.", , "Atención"
        End If
        'Si es igual a "" entonces apretó cancelar y se finaliza
    End If
End Sub


Sub Insertar_SAC2()
    Dim j As Long
    Dim i As Long
    Dim celdaDif As String
    Dim celdaMonth As String
    Dim valorVto As String
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim rango As Range
    Dim temporal As Variant
    Dim wsSAC As Excel.Worksheet
    Dim nFilasSAC As Long
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "SAC"
    Application.DisplayAlerts = True
    Set wsSAC = Worksheets("SAC")
    
    Sheets("Hoja1").Select
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count
    
    Range("1:1").Copy
    wsSAC.Range("1:1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    nFilasSAC = 1
    
    j = 2
    Do While Not IsEmpty(Cells(j, 5))
        'Identificar y agregar filas cuando corresponda para el SAC
        If Cells(j, 18).Value = 42185 Or Cells(j, 18).Value = 42551 Or Cells(j, 18).Value = 42369 Then
            
            nFilasSAC = nFilasSAC + 1
            wsSAC.Rows(nFilasSAC).Insert
            'Copio la fila completa en la nueva
            For i = 1 To nColumnas
                wsSAC.Cells(nFilasSAC, i).Value = Cells(j, i).Value
            Next i
            
            'Modifica los valores que se deben cambiar
            wsSAC.Cells(nFilasSAC, 7).Value = 316
            wsSAC.Cells(nFilasSAC, 11).Value = Cells(j, 11).Value / 2
        End If
        j = j + 1
    Loop
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub


Sub Comparar_SAC()
    Dim contenido As String
    Dim wbContenido As Workbook, _
        wsResultado As Excel.Worksheet, _
        wsContenido As Excel.Worksheet
    Dim nFilas As Double
    Dim nColumnas As Integer
    Dim rango As Range
    Dim nFilasCont As Double
    Dim rangoCont As Range
    Dim rangoTemp As String
    Dim i As Integer
    Dim j As Long
    Dim m As Integer
    Dim valorJur As Integer
    Dim valorDoc As String
    Dim valorFecha As String
    Dim resultado As Range
    Dim celdaDoc As String
    Dim temp As String
    Dim nFilasResultado As Integer
    Dim banderaSAC As Boolean

    
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
    'Agrega las nuevas hojas
    Worksheets.Add
    ActiveSheet.Name = "Resultados"
    Application.DisplayAlerts = True
    
    Set wsResultado = Worksheets("Resultados")
    Set wsContenido = wbContenido.Worksheets("Hoja1")
    
    'Regresa el control a la hoja de origen
    Sheets("SAC").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    
    'Fila del encabezado Resultado
    nFilasResultado = 1
    wsResultado.Cells(nFilasResultado, 1).Value = "JurId"
    wsResultado.Cells(nFilasResultado, 2).Value = "Doc"
    wsResultado.Cells(nFilasResultado, 3).Value = "Nombres"
    wsResultado.Cells(nFilasResultado, 4).Value = "Couc"
    wsResultado.Cells(nFilasResultado, 5).Value = "Importe"
    wsResultado.Cells(nFilasResultado, 6).Value = "Vto"
    wsResultado.Cells(nFilasResultado, 7).Value = "Importe Pagado"
    wsResultado.Cells(nFilasResultado, 8).Value = "Diferencia"
    wsResultado.Cells(nFilasResultado, 9).Value = "Observación"
    nFilasResultado = 2
    
    For i = 2 To nFilas
        valorJur = Cells(i, 12).Value
        valorDoc = Cells(i, 5).Value
        valorFecha = Cells(i, 18).Value
        banderaSAC = True
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
            If wsContenido.Cells(j, 12).Value = valorJur Then
                Do
                    If wsContenido.Cells(j, 18).Value = valorFecha Then
                        banderaSAC = False
                        'Copio a Resultado
                        wsResultado.Cells(nFilasResultado, 1).Value = Cells(i, 12).Value
                        wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 5).Value
                        wsResultado.Cells(nFilasResultado, 3).Value = Cells(i, 6).Value
                        wsResultado.Cells(nFilasResultado, 4).Value = Cells(i, 7).Value
                        wsResultado.Cells(nFilasResultado, 5).Value = Cells(i, 11).Value
                        wsResultado.Cells(nFilasResultado, 6).Value = Cells(i, 18).Value
                        wsResultado.Cells(nFilasResultado, 7).Value = wsContenido.Cells(j, 11).Value
                        wsResultado.Cells(nFilasResultado, 8).Value = Cells(i, 11).Value - wsContenido.Cells(j, 11).Value
                        wsResultado.Cells(nFilasResultado, 9).Value = "Pagado"
                        
                        If wsResultado.Cells(nFilasResultado, 8).Value = 0 Then
                            wsResultado.Cells(nFilasResultado, 10).Value = "Sin deuda"
                        End If
                        
                        'Aumento el valor del contador de las filas
                        nFilasResultado = nFilasResultado + 1
                    End If
                    j = j + 1
                Loop While valorDoc = wsContenido.Cells(j, 5).Value And banderaSAC
            End If
        End If
        If banderaSAC Then
            wsResultado.Cells(nFilasResultado, 1).Value = Cells(i, 12).Value
            wsResultado.Cells(nFilasResultado, 2).Value = Cells(i, 5).Value
            wsResultado.Cells(nFilasResultado, 3).Value = Cells(i, 6).Value
            wsResultado.Cells(nFilasResultado, 4).Value = Cells(i, 7).Value
            wsResultado.Cells(nFilasResultado, 5).Value = Cells(i, 11).Value
            wsResultado.Cells(nFilasResultado, 6).Value = Cells(i, 18).Value
            wsResultado.Cells(nFilasResultado, 9).Value = "NO Pagado"
            'Aumento el valor del contador de las filas
            nFilasResultado = nFilasResultado + 1
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


Sub Control_SAC_Residente()
    Dim contenido As String
    Dim wbContenido As Workbook, _
        wsResultado As Excel.Worksheet, _
        wsContenido As Excel.Worksheet
    Dim nFilas As Double
    Dim nColumnas As Integer
    Dim rango As Range
    Dim nFilasCont As Double
    Dim rangoCont As Range
    Dim rangoTemp As String
    Dim i As Long
    Dim j As Long
    Dim m As Integer
    Dim valorJur As Integer
    Dim valorDoc As String
    Dim valorFecha As String
    Dim resultado As Range
    Dim celdaDoc As String
    Dim temp As String
    Dim nFilasResultado As Integer
    Dim banderaSAC As Boolean

    
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
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    
    For i = 3 To 60
        valorDoc = Cells(i, 3).Value
        'Busca en el otro archivo
        rangoTemp = "C2:C" & nFilasCont
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
            
            acum = 0
            For m = 12 To 17
                acum = acum + wsContenido.Cells(j, m).Value
            Next m
            Cells(i, 12).Value = acum
            Cells(i, 13).Value = acum / 12
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub




