Attribute VB_Name = "Módulo1"
Sub Ordenar_Actuación()
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim i As Integer
    Dim valorActuacion As String
    Dim valorAnio As String
    Dim valorNum As String
    
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    For i = 1 To nFilas
        valorActuacion = Cells(i, 1).Value
        temp = ""
        cont = 0
        valorAnio = ""
        valorNum = ""
        For j = 1 To Len(valorActuacion)
            If Mid(valorActuacion, j, 1) = "-" Then
                cont = cont + 1
            End If
            If cont = 2 Then
                valorAnio = Mid(valorActuacion, j + 1, 4)
                j = j + 4
            End If
            If cont = 3 Then
                j = j + 1
                Do While Mid(valorActuacion, j, 1) <> "-"
                    valorNum = valorNum & Mid(valorActuacion, j, 1)
                    j = j + 1
                Loop
                j = j - 1
            End If
        Next j
        Cells(i, nColumnas + 1).Value = valorAnio
        Cells(i, nColumnas + 2).Value = valorNum
    Next i
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub



Sub Separar_Actuacion()
    Dim i As Long
    Dim valorVto As String
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim rango As Range
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count
    
    For i = 2 To nFilas
        valorVto = Cells(i, 4).Value
        For m = 1 To Len(valorVto)
            If Mid(valorVto, m, 1) = "-" Then
                Cells(i, 2).Value = Mid(valorVto, m + 1, 4)
                m = m + 6
                temp = ""
                Do While (Mid(valorVto, m, 1) <> "-" And m < 30)
                    temp = temp & Mid(valorVto, m, 1)
                    m = m + 1
                Loop
                Cells(i, 3).Value = temp
                If m > 28 Then
                    Cells(i, 3).Value = ""
                End If
                m = m + 5
            End If
        Next m
    Next i
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

