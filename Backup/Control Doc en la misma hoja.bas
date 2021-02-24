Attribute VB_Name = "Módulo11"
Sub Controlar_Doc()
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim i As Integer
    
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    For i = 2 To (nFilas - 1)
        valorDoc = Cells(i, 3).Value
        For j = (i + 1) To nFilas
            If Cells(j, 3).Value = valorDoc Then
                Cells(j, 3).Interior.Color = RGB(153, 196, 195)
                Cells(j, nColumnas + 1).Value = "Repetido"
            End If
        Next j
    Next i
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub


Sub Controlar_Doc2()
    Dim nFilas As Double
    Dim i As Integer
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    For i = 2 To (nFilas - 1)
        valorDoc = Cells(i, 5).Value
        valorVto = Cells(i, 12).Value
        valorCouc = Cells(i, 8).Value
        
        For j = (i + 1) To nFilas
            If Cells(j, 5).Value = valorDoc And valorVto = Cells(j, 12).Value Then
                If valorCouc = Cells(j, 8).Value Then
                    temp1 = i & ":" & i
                    Rows(temp1).Interior.Color = RGB(255, 0, 127)
                    temp2 = j & ":" & j
                    Rows(temp2).Interior.Color = RGB(102, 255, 255)
                End If
            End If
        Next j
    Next i
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Eliminar_Repetidos()
    Dim nFilas As Double
    Dim i As Integer
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    i = 2
    Do While i < nFilas
        valorDoc = Cells(i, 5).Value
        'Busca en el otro archivo
        rangoTemp = "E" & (i + 1) & ":E" & nFilas
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
            
            'Eliminar la fila y actualizar número
            Rows(j).Delete
            nFilas = nFilas - 1
            i = i - 1
        End If
        i = i + 1
    Loop
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub
