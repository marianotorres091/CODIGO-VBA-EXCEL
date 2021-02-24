Attribute VB_Name = "Módulo12"
Sub Carga()
    Dim concepto As String
    Dim ptaId As String
    Dim jurId As Integer
    Dim vencimiento As String
    Dim reajuste As String
    Dim unidad As String
    Dim actuacion As String
    Dim rangoNombre As String
    Dim rangoDocumento As String
    Dim rangoTotal As String
    Dim rangoEscalafon As String
    Dim wbDestino As Workbook, _
        wsDestino As Excel.Worksheet
    Dim rangoDest As Range
    Dim nFilasDest As Double
    Dim rango As Integer
    
    
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
    Sheets(1).Select

    'Calcular el número de filas de la hoja Destino
    Set rangoDest = wsDestino.UsedRange
    nFilasDest = rangoDest.Rows.Count
    
    'Fila del encabezado de Destino
    If nFilasDest < 2 Then
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
    
    End If
    
    nFilasDest = nFilasDest + 1
    
    jurId = InputBox("Ingrese el número de la Juridiscción:", "Jurisdicción", "1")
    ptaId = InputBox("Ingrese el rango o valor de PtaId:", "PtaId", "0")
    concepto = InputBox("Ingrese el rango o valor del Concepto:", "Concepto", "233")
    vencimiento = InputBox("Ingrese el rango o mes y año de Vencimiento:", "Vencimiento", "5/2010")
    rangoNombre = InputBox("Ingrese el rango de los Apellidos y Nombres:", "Nombre y Apellido", "A1:A80")
    rangoEscalafon = InputBox("Ingrese el rango o valor de los Escalafones:", "Escalafón", rangoNombre)
    rangoDocumento = InputBox("Ingrese el rango de los Documentos:", "Documento", rangoNombre)
    rangoTotal = InputBox("Ingrese el rango de los Totales:", "Total", rangoNombre)
    reajuste = InputBox("Ingrese el rango o valor de los Reajustes:", "Reajuste", rangoNombre)
    unidad = InputBox("Ingrese el rango o valor de las Unidades:", "Unidades", rangoNombre)
    actuacion = InputBox("Ingrese el rango o valor de las Actuaciones. Si no existen inserte 0:", "Unidades", rangoNombre)
    
    'Determina el rango de lo que se copia
    tempA = ""
    tempB = ""
    m = 2
    Do
        If IsNumeric(Mid(rangoNombre, m, 1)) Then
            tempA = tempA & Mid(rangoNombre, m, 1)
        End If
        m = m + 1
    Loop While Mid(rangoNombre, m, 1) <> ":"
    For i = (m + 1) To Len(rangoNombre)
        If IsNumeric(Mid(rangoNombre, i, 1)) Then
            tempB = tempB & Mid(rangoNombre, i, 1)
        End If
    Next i
    rango = tempB - tempA
        
    'Copia y pega los rangos completos de cada columna
    'Nombre y Apellido
    Range(rangoNombre).Copy
    temp = "G" & nFilasDest & ":G" & (nFilasDest + rango)
    wsDestino.Range(temp).PasteSpecial xlPasteValues
    'Documento
    Range(rangoDocumento).Copy
    temp = "E" & nFilasDest & ":E" & (nFilasDest + rango)
    wsDestino.Range(temp).PasteSpecial xlPasteValues
    'Importe
    Range(rangoTotal).Copy
    temp = "K" & nFilasDest & ":K" & (nFilasDest + rango)
    wsDestino.Range(temp).PasteSpecial xlPasteValues
    'Pega todo lo restante
    For i = nFilasDest To (nFilasDest + rango)
        wsDestino.Cells(i, 12).Value = Format(vencimiento, "MYYYY")
        wsDestino.Cells(i, 8).Value = concepto
        wsDestino.Cells(i, 3).Value = rangoEscalafon
        wsDestino.Cells(i, 1).Value = ptaId
        wsDestino.Cells(i, 2).Value = jurId
        wsDestino.Cells(i, 4).Value = 0
        wsDestino.Cells(i, 6).Value = 0
        wsDestino.Cells(i, 9).Value = reajuste
        wsDestino.Cells(i, 10).Value = unidad
    Next i
    'PtaId
    If Len(ptaId) > 1 Then
        Range(ptaId).Copy
        temp = "A" & nFilasDest & ":A" & (nFilasDest + rango)
        wsDestino.Range(temp).PasteSpecial xlPasteValues
    End If
    'EscId
    If Len(rangoEscalafon) > 2 Then
        Range(rangoEscalafon).Copy
        temp = "C" & nFilasDest & ":C" & (nFilasDest + rango)
        wsDestino.Range(temp).PasteSpecial xlPasteValues
    End If
    'Reajuste
    If Len(reajuste) > 2 Then
        Range(reajuste).Copy
        temp = "I" & nFilasDest & ":I" & (nFilasDest + rango)
        wsDestino.Range(temp).PasteSpecial xlPasteValues
    End If
    'Unidades
    If Len(unidad) > 1 Then
        j = tempA
        For i = nFilasDest To (nFilasDest + rango)
            tempCelda = Mid(unidad, 1, 1) & j
            porcentaje = Range(tempCelda).Value
            wsDestino.Cells(i, 10).Value = 100 * porcentaje
            j = j + 1
        Next i
    End If
    'Vencimiento
    If Mid(vencimiento, 3, 1) = ":" Or Mid(vencimiento, 4, 1) = ":" Then
        Range(vencimiento).Copy
        temp = "L" & nFilasDest & ":L" & (nFilasDest + rango)
        wsDestino.Range(temp).PasteSpecial xlPasteValues
    End If
    'Concepto
    If Len(concepto) > 3 Then
        Range(concepto).Copy
        temp = "H" & nFilasDest & ":H" & (nFilasDest + rango)
        wsDestino.Range(temp).PasteSpecial xlPasteValues
    End If
    'Actuación
    If actuacion <> "0" Then
        wsDestino.Cells(1, 13).Value = "Actuación"
        Range(actuacion).Copy
        temp = "M" & nFilasDest & ":M" & (nFilasDest + rango)
        wsDestino.Range(temp).PasteSpecial xlPasteValues
    End If
    
    Application.CutCopyMode = False
    wsDestino.Application.CutCopyMode = False
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "Ha ocurrido un error. Comprube que los datos ingresados son correctos.", , "Error"
End Sub


Sub Eliminar_Importe()
    Dim nFilas As Double
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    For i = 2 To nFilas
        If Cells(i, 11).Value = 0 Or Cells(i, 11).Value = "" Then
            Rows(i).Delete
            i = i - 1
            nFilas = nFilas - 1
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Corregir_Unidad()
    Dim nFilas As Double
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    For i = 2 To nFilas
        temp = Cells(i, 10).Value * 100
        Cells(i, 10).Value = temp
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub


Sub SAC_Separado()
    Dim concepto As String
    Dim i As Integer
    Dim vencimiento As Integer
    Dim nombre As Integer
    Dim documento As Integer
    Dim total As Integer
    Dim wbDestino As Workbook, _
        wsDestino As Excel.Worksheet
    Dim rangoDest As Range
    Dim nFilasDest As Double
    Dim nFilas As Double
    Dim rango As Range
    
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
    Sheets(1).Select
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count

    'Calcular el número de filas de la hoja Destino
    Set rangoDest = wsDestino.UsedRange
    nFilasDest = rangoDest.Rows.Count

    vencimiento = InputBox("Ingrese el número de columna de mes y año de Vencimiento:", "Vencimiento", "5")
    nombre = InputBox("Ingrese el número de columna de Apellidos y Nombres:", "Nombre y Apellido", "6")
    documento = InputBox("Ingrese el número de columna de Documentos:", "Documento", "4")
    total = InputBox("Ingrese el número de columna del SAC:", "SAC", "7")
    
    For i = 2 To nFilas
        If Cells(i, total).Value <> "" Then
            'Copiar fila
            wsDestino.Cells(nFilasDest, 5).Value = Cells(i, documento).Value
            wsDestino.Cells(nFilasDest, 12).Value = Cells(i, vencimiento).Value
            wsDestino.Cells(nFilasDest, 11).Value = Cells(i, total).Value
            wsDestino.Cells(nFilasDest, 7).Value = Cells(i, nombre).Value
            wsDestino.Cells(nFilasDest, 8).Value = 316
            nFilasDest = nFilasDest + 1
        End If
    Next i
    
    Application.CutCopyMode = False
    wsDestino.Application.CutCopyMode = False
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "Ha ocurrido un error. Comprube que los datos ingresados son correctos.", , "Error"
End Sub

