Attribute VB_Name = "Módulo1"
Sub Filtrar_Docuementos()
    Dim contenido As String
    Dim wbContenido As Workbook, _
        wsResultado As Excel.Worksheet, _
        wsError As Excel.Worksheet, _
        wsContenido As Excel.Worksheet
    Dim nFilas As Double
    Dim rango As Range
    Dim nFilasCont As Double
    Dim rangoCont As Range
    Dim rangoError As Range
    Dim rangoTemp As String
    Dim i As Integer
    Dim j As Long
    Dim m As Integer
    Dim valorJur As Integer
    Dim valorDoc As String
    Dim resultado As Range
    Dim celdaDoc As String
    Dim temp As String
    Dim nFilasResultado As Integer
    Dim nFilasError As Integer
    Dim nColumnasError As Integer
    Dim bandera As Boolean
    
    
    'Indicar el libro de excel CONTENIDO y control de errores
    contenido = InputBox("Ingrese el nombre del archivo:", "Abrir", "probando.xlsx")
    If contenido <> "" Then
        On Error GoTo ControlErrorOpen
        Set wbContenido = Workbooks.Open(ActiveWorkbook.Path & "\" & contenido)
    Else
        Exit Sub
    End If
    
    'Activar este libro
    ThisWorkbook.Activate
    
    'Borra las hojas destino si existen
    Application.DisplayAlerts = False
    'Sheets("Resultados").Delete
    'Sheets("Errores").Delete
    
    'Agrega las nuevas hojas
    Worksheets.Add
    ActiveSheet.Name = "Resultados"
    Worksheets.Add
    ActiveSheet.Name = "Errores"
    Application.DisplayAlerts = True
    
    Set wsResultado = Worksheets("Resultados")
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
    
    'Fila del encabezado Error
    wsContenido.Range("A1").Copy
    wsResultado.Range("A1:L1").PasteSpecial Paste:=xlFormats
    Application.CutCopyMode = False
    nFilasResultado = 1
    wsResultado.Cells(nFilasResultado, 1).Value = "PtaId"
    wsResultado.Cells(nFilasResultado, 2).Value = "JurId"
    wsResultado.Cells(nFilasResultado, 3).Value = "EscId"
    wsResultado.Cells(nFilasResultado, 4).Value = "Pref"
    wsResultado.Cells(nFilasResultado, 5).Value = "Doc"
    wsResultado.Cells(nFilasResultado, 6).Value = "Digito"
    wsResultado.Cells(nFilasResultado, 7).Value = "Nombres"
    wsResultado.Cells(nFilasResultado, 8).Value = "Couc"
    wsResultado.Cells(nFilasResultado, 9).Value = "Reajuste"
    wsResultado.Cells(nFilasResultado, 10).Value = "Unidades"
    wsResultado.Cells(nFilasResultado, 11).Value = "Importe"
    wsResultado.Cells(nFilasResultado, 12).Value = "Vto"
    'wsResultado.Range("A2:D2") = ["apokd","amñfmdf","afosmf","304k4"]
    nFilasResultado = 2
    
    'SE PUEDE PROBAR SI DEJA METER UN FILTRO DENTRO DE OTRO. PROBANDO POR EJEMPLO EL VALOR DE RANGO. O ALGO QUE DEJE
    
    For i = 2 To nFilas
        valorJur = Cells(i, 1).Value
        valorDoc = Cells(i, 2).Value
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
        bandera = True
        If Not resultado Is Nothing Then
                primerResultado = resultado.Address
            Do
                'Se obtiene el valor de j
                celdaDoc = resultado.Address
                tempDoc = ""
                For m = 1 To Len(celdaDoc)
                    If IsNumeric(Mid(celdaDoc, m, 1)) Then
                        tempDoc = tempDoc & Mid(celdaDoc, m, 1)
                    End If
                Next m
                j = tempDoc
                'Identifica y agrega filas cuando corresponda
                If wsContenido.Cells(j, 1).Value = valorJur Then
                    'Copio y pego la fila de Contenido a Resultado
                    wsResultado.Cells(nFilasResultado, 1).Value = wsContenido.Cells(j, 2).Value
                    wsResultado.Cells(nFilasResultado, 2).Value = wsContenido.Cells(j, 1).Value
                    wsResultado.Cells(nFilasResultado, 3).Value = wsContenido.Cells(j, 7).Value
                    wsResultado.Cells(nFilasResultado, 4).Value = wsContenido.Cells(j, 3).Value
                    wsResultado.Cells(nFilasResultado, 5).Value = wsContenido.Cells(j, 4).Value
                    wsResultado.Cells(nFilasResultado, 6).Value = wsContenido.Cells(j, 5).Value
                    wsResultado.Cells(nFilasResultado, 7).Value = wsContenido.Cells(j, 6).Value
                    wsResultado.Cells(nFilasResultado, 8).Value = wsContenido.Cells(j, 15).Value
                    wsResultado.Cells(nFilasResultado, 9).Value = 1
                    wsResultado.Cells(nFilasResultado, 10).Value = 0
                    wsResultado.Cells(nFilasResultado, 11).Value = wsContenido.Cells(j, 19).Value
                    wsResultado.Cells(nFilasResultado, 12).Value = wsContenido.Cells(j, 20).Value
                    nFilasResultado = nFilasResultado + 1
                    'Actualizo la bandera porque se econtró
                    bandera = False
                End If
                Set resultado = wsContenido.Range(rangoTemp).FindNext(resultado)
            Loop While Not resultado Is Nothing And primerResultado <> resultado.Address
        Else
            'No se encontró el documento
            'Copio la fila de Origen
            temp = i & ":" & i
            Range(temp).Copy
            'Pegar y actualizar num de filas. Agregando el msj correspondiente
            temp = nFilasError & ":" & nFilasError
            wsError.Range(temp).PasteSpecial xlPasteAll
            Application.CutCopyMode = False
            wsError.Cells(nFilasError, nColumnasError).Value = "No se encontró el Documento."
            nFilasError = nFilasError + 1
            bandera = False
        End If
        If bandera Then
            'No se encontró el documento en la jurisdicción
            'Copio la fila de Origen
            temp = i & ":" & i
            Range(temp).Copy
            'Pegar y actualizar num de filas. Agregando el msj correspondiente
            temp = nFilasError & ":" & nFilasError
            wsError.Range(temp).PasteSpecial xlPasteAll
            Application.CutCopyMode = False
            wsError.Cells(nFilasError, nColumnasError).Value = "No se encontró el Documento en la Jurisdicción indicada."
            nFilasError = nFilasError + 1
        End If
    Next i
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub
