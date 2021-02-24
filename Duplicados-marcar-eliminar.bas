Attribute VB_Name = "Módulo2"
Sub Eliminar_Duplicados()
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim i As Integer
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    For i = 3 To nFilas
        bandera = True
        For j = 1 To nColumnas
            If Cells(i, j).Value <> Cells(i - 1, j).Value Then
                bandera = False
            End If
        Next j
        If bandera Then
            Rows(i).Delete
            nFilas = nFilas - 1
            i = i - 1
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Marcar_Duplicados()
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim i As Integer
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    Cells(1, nColumnas + 1).Value = "Repetidos"
    For i = 3 To nFilas
        bandera = True
        For j = 1 To nColumnas
            If Cells(i, j).Value <> Cells(i - 1, j).Value Then
                bandera = False
            End If
        Next j
        If bandera Then
            Cells(i, nColumnas + 1).Value = "Repetido"
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Eliminar_Duplicados_2()
    Dim nFilas As Long
    Dim nFilasCuotas As Long
    Dim i As Long
    Dim j As Long
    Dim wsCuota1 As Excel.Worksheet
    Dim resultado As Range
    
    Set wsCuota1 = Worksheets("Cuota Pagada")
    
    'Regresa el control a la hoja de origen
    Sheets("Restante").Select
    
    Set rangoCont = wsCuota1.UsedRange
    nFilasCuotas = rangoCont.Rows.Count
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    nColumnas = nColumnas - 1
    
    For i = 2 To nFilasCuotas
        valorDoc = wsCuota1.Cells(i, 5).Value
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
            j = tempDoc - 1
            Do
                bandera = True
                For m = 1 To nColumnas
                    If Cells(j, m).Value <> wsCuota1.Cells(i, m).Value Then
                        bandera = False
                    End If
                Next m
                If bandera Then
                    Rows(j).Delete
                    nFilas = nFilas - 1
                    wsCuota1.Cells(i, nColumnas + 1).Value = "Eliminado"
                End If
                j = j + 1
            Loop While (valorDoc = Cells(j, 5).Value)
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub


