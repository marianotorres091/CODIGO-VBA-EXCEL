Attribute VB_Name = "Módulo1"
Sub Filtrar_Agentes()
    Dim wsResultado As Excel.Worksheet
    Dim rango As Range
    Dim nFilas As Double
    Dim nFilasResultado As Integer
    Dim importeTotal As Double
    Dim cantidad As Integer
    Dim docAnterior As String
    
    
    MsgBox "Debe estar ordenado por Mes y DNI.", , "¡Atención!"
    
    Application.DisplayAlerts = False
    'Agrega las nuevas hojas
    Worksheets.Add
    ActiveSheet.Name = "Resultado"
    Application.DisplayAlerts = True
    
    Set wsResultado = Worksheets("Resultado")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    'Fila del encabezado Resultado
    nFilasResultado = 2
    wsResultado.Cells(1, 1).Value = "Año"
    wsResultado.Cells(1, 2).Value = "Mes"
    wsResultado.Cells(1, 3).Value = "JurId"
    wsResultado.Cells(1, 4).Value = "Documento"
    wsResultado.Cells(1, 5).Value = "Nombre y Apellido"
    wsResultado.Cells(1, 6).Value = "Concepto"
    wsResultado.Cells(1, 7).Value = "Cantidad"
    wsResultado.Cells(1, 8).Value = "Importe Total"
    wsResultado.Range("1:1").Font.Bold = True
    wsResultado.Range("1:1").HorizontalAlignment = xlCenter
    
    'Tratar al primero
    cantidad = 1
    importeTotal = Cells(2, 7).Value
    docAnterior = Cells(2, 12).Value
    For i = 3 To nFilas
        'NO SE COMPARA LOS MESES, PORQUE EL ARCHIVO TIENE UN MES UNICAMENTE
        If docAnterior = Cells(i, 12).Value Then
            'Acumulo
            cantidad = cantidad + 1
            importeTotal = importeTotal + Cells(i, 7).Value
        Else
            'Imprimir
            wsResultado.Cells(nFilasResultado, 1).Value = Cells(i - 1, 1).Value
            wsResultado.Cells(nFilasResultado, 2).Value = Cells(i - 1, 2).Value
            wsResultado.Cells(nFilasResultado, 3).Value = Cells(i - 1, 8).Value
            wsResultado.Cells(nFilasResultado, 4).Value = Cells(i - 1, 12).Value
            wsResultado.Cells(nFilasResultado, 5).Value = Cells(i - 1, 14).Value
            wsResultado.Cells(nFilasResultado, 6).Value = Cells(i - 1, 4).Value
            wsResultado.Cells(nFilasResultado, 7).Value = cantidad
            wsResultado.Cells(nFilasResultado, 8).Value = importeTotal
            nFilasResultado = nFilasResultado + 1
            'Tratar nuevo
            cantidad = 1
            importeTotal = Cells(i, 7).Value
            docAnterior = Cells(i, 12).Value
        End If
    Next i
    'Imprimir lo último
    wsResultado.Cells(nFilasResultado, 1).Value = Cells(i - 1, 1).Value
    wsResultado.Cells(nFilasResultado, 2).Value = Cells(i - 1, 2).Value
    wsResultado.Cells(nFilasResultado, 3).Value = Cells(i - 1, 8).Value
    wsResultado.Cells(nFilasResultado, 4).Value = Cells(i - 1, 12).Value
    wsResultado.Cells(nFilasResultado, 5).Value = Cells(i - 1, 14).Value
    wsResultado.Cells(nFilasResultado, 6).Value = Cells(i - 1, 4).Value
    wsResultado.Cells(nFilasResultado, 7).Value = cantidad
    wsResultado.Cells(nFilasResultado, 8).Value = importeTotal
            
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    
End Sub

Sub Eliminar_Unos()
    Dim i As Double
    
    i = 2
    Do While Cells(i, 4).Value <> ""
        If Cells(i, 7).Value = 1 Then
            Rows(i).Delete
        Else
            i = i + 1
        End If
    Loop
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub


Sub Eliminar_Agentes()
    Dim contenido As String
    Dim wbContenido As Workbook, _
        wsError As Excel.Worksheet, _
        wsContenido As Excel.Worksheet
    Dim nFilas As Double
    Dim primerFila As Double
    Dim rango As Range
    Dim nFilasCont As Double
    Dim nColumnasCont As Double
    Dim rangoCont As Range
    Dim rangoError As Range
    Dim rangoTemp As String
    Dim i As Long
    Dim j As Long
    Dim m As Long
    Dim valorJur As Integer
    Dim valorDoc As String
    Dim resultado As Range
    Dim celdaDoc As String
    Dim nFilasError As Integer
    Dim nColumnasError As Integer
    Dim bandera As Boolean
    Dim totalImporte As Double
    Dim contador As Integer

    
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
    ActiveSheet.Name = "Errores"
    Application.DisplayAlerts = True
    
    
    Set wsError = Worksheets("Errores")
    Set wsContenido = wbContenido.Worksheets("Detalle x Agente")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    nColumnasCont = rangoCont.Columns.Count
    
    'Fila del encabezado Error
    Range("1:1").Copy
    wsError.Range("1:1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    nFilasError = 2
    Set rangoError = wsError.UsedRange
    nColumnasError = rangoError.Columns.Count
    nColumnasError = nColumnasError + 1
    wsError.Range("A1").Copy
    wsError.Cells(1, nColumnasError).PasteSpecial Paste:=xlFormats
    Application.CutCopyMode = False
    wsError.Cells(1, nColumnasError).Value = "Mensaje"
    wsError.Columns(nColumnasError).ColumnWidth = 52
    
    
    For i = 2 To nFilas
        valorJur = Cells(i, 2).Value
        valorDoc = Cells(i, 5).Value
        totalImporte = 0
        contador = 0
        bandera = False
        'Busca en el otro archivo
        rangoTemp = "D2:D" & nFilasCont
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
            If wsContenido.Cells(j, 1).Value = valorJur Then
                'retrocedo hasta encontrar el primero
                bandera = True
                Do While bandera
                    If wsContenido.Cells(j - 1, 4).Value = valorDoc Then
                        j = j - 1
                    Else
                        bandera = False
                    End If
                Loop
                'Estoy parado en la primer celda del Doc
                primerFila = j
                Do
                    If wsContenido.Cells(j, 15).Value = 233 Then
                        totalImporte = totalImporte + wsContenido.Cells(j, 19).Value
                        contador = contador + 1
                    End If
                    'Marcar las filas
                    wsContenido.Cells(j, nColumnasCont + 1).Value = "Cobrado"
                    j = j + 1
                Loop While valorDoc = wsContenido.Cells(j, 4).Value
                
                bandera = False
                If Cells(i, 11).Value <> totalImporte Then
                    wsError.Cells(nFilasError, nColumnasError).Value = "Diferencia de Importe Total:"
                    wsError.Cells(nFilasError, nColumnasError + 1).Value = totalImporte - Cells(i, 8).Value
                    bandera = True
                End If
                'Marcar las filas
                'For m = primerFila To j
                 '   wsContenido.Cells(m, nColumnasCont + 1).Value = "Cobrado"
                'Next m
            Else
                'No se encontró el documento en la jurisdicción
                wsError.Cells(nFilasError, nColumnasError).Value = "No se encontró el Documento en la Jurisdicción indicada. Está en la " & wsContenido.Cells(j, 1).Value
                bandera = True
            End If
        Else
            'No se encontró el documento
            wsError.Cells(nFilasError, nColumnasError).Value = "No se encontró el Documento."
            bandera = True
        End If
        If bandera Then
            wsError.Cells(nFilasError, 1).Value = Cells(i, 1).Value
            wsError.Cells(nFilasError, 2).Value = Cells(i, 2).Value
            wsError.Cells(nFilasError, 3).Value = Cells(i, 3).Value
            wsError.Cells(nFilasError, 4).Value = Cells(i, 4).Value
            wsError.Cells(nFilasError, 5).Value = Cells(i, 5).Value
            wsError.Cells(nFilasError, 6).Value = Cells(i, 6).Value
            wsError.Cells(nFilasError, 7).Value = Cells(i, 7).Value
            wsError.Cells(nFilasError, 8).Value = Cells(i, 8).Value
            nFilasError = nFilasError + 1
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


Sub Juntar_Duplicados()
    Dim nFilas As Double
    Dim i As Integer
    
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    For i = 2 To (nFilas - 1)
        If Cells(i, 4).Value = Cells(i + 1, 4).Value And Cells(i, 4).Value <> "" Then
            nuevoImporte = Cells(i, 8).Value + Cells(i + 1, 8).Value
            nuevaCantidad = Cells(i, 7).Value + Cells(i + 1, 7).Value
            Cells(i, 7).Value = nuevaCantidad
            Cells(i, 8).Value = nuevoImporte
            Rows(i + 1).Delete
            nFilas = nFilas - 1
            i = i - 1
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub


Sub Eliminar_Cobrados()
    Dim i As Long
    Dim nFilas As Double
    Dim nColumnas As Long
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    For i = 2 To nFilas
        If Cells(i, nColumnas).Value = "Cobrado" Then
            Rows(i).Delete
            i = i - 1
            nFilas = nFilas - 1
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub


Sub Eliminas_Bajos()
    Dim i As Long
    Dim nFilas As Double
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    For i = 2 To nFilas
        If Cells(i, 10).Value <> "" Then
            If Cells(i, 10).Value <= 100 Then
                Rows(i).Delete
                i = i - 1
                nFilas = nFilas - 1
            End If
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub
