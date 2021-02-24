Attribute VB_Name = "Módulo12"
Sub Corregir_Vto()
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim temp As String
    
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    nColumnas = nColumnas + 1
    
    'Agregar columna VTO
    Cells(1, nColumnas).Value = "Vto"
    Range("A1").Copy
    Cells(1, nColumnas).PasteSpecial Paste:=xlFormats
    Application.CutCopyMode = False
    
    'columna 7
    For i = 2 To nFilas
        temp = Format(Cells(i, 7).Value, "MYYYY")
        Cells(i, nColumnas).Value = temp
    Next i
    
    'Mostrar msj para confirmar
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
