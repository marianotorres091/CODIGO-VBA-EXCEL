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


Sub Separar_Actuacion_2()
    Dim i As Long
    Dim valorVto As String
    Dim nFilas As Double
    Dim nColumnas As Double
    Dim rango As Range
    
    'Calcular el número de filas y columnas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nColumnas = rango.Columns.Count
    nFilas = rango.Rows.Count
    
    For i = 6 To nFilas
        valorVto = Cells(i, 1).Value
        m = 1
        Do While m < Len(valorVto)
            If IsNumeric(Mid(valorVto, m, 1)) Then
                Cells(i, nColumnas + 1).Value = Cells(i, nColumnas + 1).Value & Mid(valorVto, m, 1)
            End If
            If Mid(valorVto, m, 1) = "-" Then
                Cells(i, nColumnas + 2).Value = Mid(valorVto, m + 1, 4)
                m = m + 6
                temp = ""
                Do While (Mid(valorVto, m, 1) <> "-" And m < 30)
                    temp = temp & Mid(valorVto, m, 1)
                    m = m + 1
                Loop
                Cells(i, nColumnas + 3).Value = temp
                If m > 28 Then
                    Cells(i, nColumnas + 3).Value = ""
                End If
                m = m + 5
            End If
            m = m + 1
        Loop
    Next i
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Comparar_Archivos()
    Dim contenido As String
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet
    Dim nFilas As Double
    Dim nColumnas As Integer
    Dim rango As Range
    Dim nFilasCont As Double
    Dim nColumnasCont As Integer
    Dim rangoCont As Range
    Dim i As Long
    Dim j As Long
    Dim jur As Integer
    
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
    
    'Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    nColumnasCont = rangoCont.Columns.Count
    
    Cells(1, nColumnas + 1).Value = "PROCESO"
    
    For i = 2 To nFilas
        For j = 8 To nFilasCont
            If Cells(i, 5).Value = wsContenido.Cells(j, 11).Value Then
                jur = Mid(Cells(i, 3).Value, 2, Len(Cells(i, 3).Value) - 1)
                If Cells(i, 4).Value = wsContenido.Cells(j, 10).Value Then
                    Cells(i, nColumnas + 1).Value = "EN GESTIÓN"
                    wsContenido.Cells(j, nColumnasCont + 1).Value = "ENCONTRADA"
                    If jur <> wsContenido.Cells(j, 9).Value Then
                        Cells(i, nColumnas + 1).Value = "EN GESTIÓN. VER JUR"
                    End If
                    j = nFilasCont
                End If
            End If
        Next j
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub
