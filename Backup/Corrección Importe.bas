Attribute VB_Name = "Módulo1"
Sub Corregir_Importe()
    Dim columnaImporte As Integer
    Dim nFilas As Double
    Dim temp As String
    Dim nuevoValor As Double
    Dim rango As Range
    
    columnaImporte = 11
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    For i = 2 To nFilas
        valorCelda = Cells(i, columnaImporte).Value
        temp = ""
        For j = 1 To Len(valorCelda)
            If Mid(valorCelda, j, 1) = "." Then
                temp = temp & ","
            Else
                temp = temp & Mid(valorCelda, j, 1)
            End If
        Next j
        nuevoValor = temp
        Cells(i, columnaImporte).Value = nuevoValor
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub


Sub Eliminar_CEROS()
    Dim columnaImporte As Integer
    Dim nFilas As Double
    Dim temp As String
    Dim nuevoValor As Double
    Dim rango As Range
    
    columnaImporte = 8
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    For i = 2 To nFilas
        If Cells(i, columnaImporte).Value <> "" Then
            If Cells(i, columnaImporte).Value < 10 And Cells(i, columnaImporte).Value > (-10) Then
                Rows(i).Delete
                i = i - 1
                nFilas = nFilas - 1
            End If
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub


Sub Nuevo_Ceic()
    Dim contenido As String
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet
    Dim nFilas As Double
    Dim nColumnas As Integer
    Dim rango As Range
    Dim nFilasCont As Double
    Dim rangoCont As Range
    Dim rangoTemp As String
    Dim i As Long
    Dim j As Long
    Dim valorDoc As String
    Dim resultado As Range
    Dim celdaDoc As String
    Dim temp As String
    Dim bandera As Boolean
    Dim totalImporte As Double

    
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
    
    Cells(1, nColumnas + 1).Value = "Importe Anterior"
    Cells(1, nColumnas + 2).Value = "Importe Nuevo"
    Cells(1, nColumnas + 3).Value = "Diferencia"

    
    For i = 2 To nFilas
        valorDoc = Cells(i, 16).Value
        'Busca en el otro archivo
        rangoTemp = "L3:L" & nFilasCont
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
            
            Cells(i, nColumnas + 1).Value = wsContenido.Cells(j, 7).Value
            
            temp = 0
            For m = (j - 2) To j
                If wsContenido.Cells(m, 13).Value = valorDoc Then
                    temp = m
                End If
                If wsContenido.Cells(m, 14).Value = valorDoc Then
                    temp = m
                End If
                If wsContenido.Cells(m, 15).Value = valorDoc Then
                    temp = m
                End If
            Next m
            
            If temp = 0 Then
                Cells(i, nColumnas + 4).Value = "ERROR - Controlar"
            Else
                Cells(i, nColumnas + 2).Value = wsContenido.Cells(temp, 7).Value
                Cells(i, nColumnas + 3).Value = Cells(i, nColumnas + 2).Value - Cells(i, nColumnas + 1).Value
            End If
        Else
            'No se encontró el documento
            Cells(i, nColumnas + 4).Value = "No se encontró el DNI"
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub
