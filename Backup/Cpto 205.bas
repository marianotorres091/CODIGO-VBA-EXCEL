Attribute VB_Name = "Módulo1"
Sub Contar_Liquidaciones_Mes()
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet
    Dim valorDoc As String
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim nColumnasCont As Long
    Dim nFilasCont As Long
    Dim i As Long
    Dim contador As Integer
    Dim mes As Integer
    Dim valorJur As Integer
    Dim anio As Integer
    Dim importeTotal As Double
    
    
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
    
    wsContenido.Cells(1, 1).Value = "JUR"
    wsContenido.Cells(1, 2).Value = "DNI"
    wsContenido.Cells(1, 3).Value = "Nombre"
    wsContenido.Cells(1, 4).Value = "Año"
    wsContenido.Cells(1, 5).Value = "Mes"
    wsContenido.Cells(1, 6).Value = "Cantidad"
    wsContenido.Cells(1, 7).Value = "Importe Total"
    wsContenido.Cells(1, 8).Value = "Último CEIC"
    wsContenido.Cells(1, 9).Value = "Observación"
    nFilasCont = 2
    
    valorDoc = Cells(2, 12).Value
    valorJur = Cells(2, 8).Value
    importeTotal = Cells(2, 7).Value
    mes = Cells(2, 2).Value
    anio = Cells(2, 1).Value
    contador = 1
    
    wsContenido.Cells(nFilasCont, 1).Value = Cells(2, 8).Value
    wsContenido.Cells(nFilasCont, 2).Value = valorDoc
    wsContenido.Cells(nFilasCont, 3).Value = Cells(2, 14).Value
    bandera = False
    
    For i = 3 To nFilas
        If valorDoc = Cells(i, 12).Value And valorJur <> Cells(i, 8).Value Then
            bandera = True
        End If
        If valorDoc = Cells(i, 12).Value And anio = Cells(i, 1).Value And mes = Cells(i, 2).Value Then
            contador = contador + 1
            If Cells(i, 6).Value = 2 Then
                importeTotal = importeTotal - Cells(i, 7).Value
            Else
                importeTotal = importeTotal + Cells(i, 7).Value
            End If
        Else
            'Copio en archivo
            wsContenido.Cells(nFilasCont, 4).Value = anio
            wsContenido.Cells(nFilasCont, 5).Value = mes
            wsContenido.Cells(nFilasCont, 6).Value = contador
            wsContenido.Cells(nFilasCont, 7).Value = importeTotal
            wsContenido.Cells(nFilasCont, 8).Value = Cells(i - 1, 15).Value
            If bandera Then
                wsContenido.Cells(nFilasCont, 9).Value = "Varias JUR"
            End If
            nFilasCont = nFilasCont + 1
            'Reinicio contadores
            valorDoc = Cells(i, 12).Value
            mes = Cells(i, 2).Value
            anio = Cells(i, 1).Value
            valorJur = Cells(i, 8).Value
            bandera = False
            'Trato linea
            contador = 1
            importeTotal = Cells(i, 7).Value
            
            wsContenido.Cells(nFilasCont, 1).Value = Cells(i, 8).Value
            wsContenido.Cells(nFilasCont, 2).Value = valorDoc
            wsContenido.Cells(nFilasCont, 3).Value = Cells(i, 14).Value
        End If
    Next i
    
    wsContenido.Cells(nFilasCont, 4).Value = anio
    wsContenido.Cells(nFilasCont, 5).Value = mes
    wsContenido.Cells(nFilasCont, 6).Value = contador
    wsContenido.Cells(nFilasCont, 7).Value = importeTotal
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


Sub Contar_Liquidaciones_Mes_JUR()
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet
    Dim valorDoc As String
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim nColumnasCont As Long
    Dim nFilasCont As Long
    Dim i As Long
    Dim contador As Integer
    Dim mes As Integer
    Dim jur As Integer
    Dim anio As Integer
    Dim importeTotal As Double
    
    
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
    
    wsContenido.Cells(1, 1).Value = "JUR"
    wsContenido.Cells(1, 2).Value = "DNI"
    wsContenido.Cells(1, 3).Value = "Nombre"
    wsContenido.Cells(1, 4).Value = "Año"
    wsContenido.Cells(1, 5).Value = "Mes"
    wsContenido.Cells(1, 6).Value = "Cantidad"
    wsContenido.Cells(1, 7).Value = "Importe Total"
    nFilasCont = 2
    
    jur = Cells(2, 8).Value
    valorDoc = Cells(2, 12).Value
    importeTotal = Cells(2, 7).Value
    mes = Cells(2, 2).Value
    anio = Cells(2, 1).Value
    contador = 1
    
    wsContenido.Cells(nFilasCont, 1).Value = jur
    wsContenido.Cells(nFilasCont, 2).Value = valorDoc
    wsContenido.Cells(nFilasCont, 3).Value = Cells(2, 14).Value
    
    For i = 3 To nFilas
        If jur = Cells(i, 8).Value And valorDoc = Cells(i, 12).Value And anio = Cells(i, 1).Value And mes = Cells(i, 2).Value Then
            contador = contador + 1
            If Cells(i, 6).Value = 2 Then
                importeTotal = importeTotal - Cells(i, 7).Value
            Else
                importeTotal = importeTotal + Cells(i, 7).Value
            End If
        Else
            'Copio en archivo
            wsContenido.Cells(nFilasCont, 4).Value = anio
            wsContenido.Cells(nFilasCont, 5).Value = mes
            wsContenido.Cells(nFilasCont, 6).Value = contador
            wsContenido.Cells(nFilasCont, 7).Value = importeTotal
            nFilasCont = nFilasCont + 1
            'Reinicio contadores
            jur = Cells(i, 8).Value
            valorDoc = Cells(i, 12).Value
            mes = Cells(i, 2).Value
            anio = Cells(i, 1).Value
            'Trato linea
            contador = 1
            importeTotal = Cells(i, 7).Value
            
            wsContenido.Cells(nFilasCont, 1).Value = jur
            wsContenido.Cells(nFilasCont, 2).Value = valorDoc
            wsContenido.Cells(nFilasCont, 3).Value = Cells(i, 14).Value
        End If
    Next i
    
    wsContenido.Cells(nFilasCont, 4).Value = anio
    wsContenido.Cells(nFilasCont, 5).Value = mes
    wsContenido.Cells(nFilasCont, 6).Value = contador
    wsContenido.Cells(nFilasCont, 7).Value = importeTotal
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


Sub Contar_Liquidaciones_JUR()
    Dim wsResultado As Excel.Worksheet
    Dim valorDoc As String
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim nColumnasCont As Long
    Dim nFilasCont As Long
    Dim i As Long
    Dim contador As Integer
    Dim jur As Integer
    Dim importeTotal As Double
    
    'Activar este libro
    ThisWorkbook.Activate

    'Agrega las nuevas hojas
    Worksheets.Add
    ActiveSheet.Name = "Resultados"
    Application.DisplayAlerts = True
    
    Set wsResultado = Worksheets("Resultados")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select

    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    wsResultado.Cells(1, 1).Value = "JUR"
    wsResultado.Cells(1, 2).Value = "DNI"
    wsResultado.Cells(1, 3).Value = "Nombre"
    wsResultado.Cells(1, 4).Value = "Cantidad"
    wsResultado.Cells(1, 5).Value = "Importe Total"
    nFilasCont = 2
    
    jur = Cells(2, 7).Value
    valorDoc = Cells(2, 8).Value
    importeTotal = Cells(2, 6).Value
    contador = 1
    
    wsResultado.Cells(nFilasCont, 1).Value = jur
    wsResultado.Cells(nFilasCont, 2).Value = valorDoc
    wsResultado.Cells(nFilasCont, 3).Value = Cells(2, 9).Value
    
    For i = 3 To nFilas
        If jur = Cells(i, 7).Value And valorDoc = Cells(i, 8).Value Then
            contador = contador + 1
            If Cells(i, 5).Value = 2 Then
                importeTotal = importeTotal - Cells(i, 6).Value
            Else
                importeTotal = importeTotal + Cells(i, 6).Value
            End If
        Else
            'Copio en archivo
            wsResultado.Cells(nFilasCont, 4).Value = contador
            wsResultado.Cells(nFilasCont, 5).Value = importeTotal
            nFilasCont = nFilasCont + 1
            'Reinicio contadores
            jur = Cells(i, 7).Value
            valorDoc = Cells(i, 8).Value
            contador = 1
            importeTotal = Cells(i, 6).Value
            
            wsResultado.Cells(nFilasCont, 1).Value = jur
            wsResultado.Cells(nFilasCont, 2).Value = valorDoc
            wsResultado.Cells(nFilasCont, 3).Value = Cells(i, 9).Value
        End If
    Next i
    
    wsResultado.Cells(nFilasCont, 4).Value = contador
    wsResultado.Cells(nFilasCont, 5).Value = importeTotal
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub



