Attribute VB_Name = "Módulo2"
Sub Cuota_8()
    Dim wsCuota1 As Excel.Worksheet
    Dim wsRestoCuotas As Excel.Worksheet
    Dim rango As Range
    Dim nFilas As Long
    Dim nFilasCuota As Long
    Dim nFilasResto As Long
    Dim importeTotal As Double
    Dim montoCuota As Double
    Dim tempPrimero As Integer
    Dim tempUltimo As Integer
    Dim dni As String
    
    
    MsgBox "Debe estar ordenado por DNI.", , "¡Atención!"
    
    Application.DisplayAlerts = False
    'Agrega las nuevas hojas
    Worksheets.Add
    ActiveSheet.Name = "Cuota 1"
    Worksheets.Add
    ActiveSheet.Name = "Cuotas Faltantes"
    Application.DisplayAlerts = True
    
    Set wsCuota1 = Worksheets("Cuota 1")
    Set wsRestoCuotas = Worksheets("Cuotas Faltantes")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    'Fila del encabezado
    nFilasCuota = 2
    nFilasResto = 2
    
    Range("1:1").Copy
    wsCuota1.Range("1:1").PasteSpecial xlPasteAll
    wsRestoCuotas.Range("1:1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    
    wsCuota1.Range("1:1").Font.Bold = True
    wsCuota1.Range("1:1").HorizontalAlignment = xlCenter
    wsRestoCuotas.Range("1:1").Font.Bold = True
    wsRestoCuotas.Range("1:1").HorizontalAlignment = xlCenter
    
    'Tratar al primero
    montoCuota = 0
    importeTotal = 0
    dni = Cells(2, 5).Value
    tempPrimero = 2
    For i = 2 To nFilas
        If dni = Cells(i, 5).Value Then
            'Acumulo
            If Cells(i, 9).Value = 2 Then
                importeTotal = importeTotal - Cells(i, 11).Value
            Else
                importeTotal = importeTotal + Cells(i, 11).Value
            End If
        Else
            'Fin DNI
            tempUltimo = i - 1
            montoCuota = importeTotal / 8
            'Cargar Cuotas
            monto = 0
            bandera = True
            For j = tempPrimero To tempUltimo
                If bandera Then
                    If Cells(j, 9).Value = 2 Then
                        monto = monto - Cells(j, 11).Value
                    Else
                        monto = monto + Cells(j, 11).Value
                    End If
                    If monto < montoCuota Then
                        temp = j & ":" & j
                        Range(temp).Copy
                        temp = nFilasCuota & ":" & nFilasCuota
                        wsCuota1.Range(temp).PasteSpecial xlPasteAll
                        nFilasCuota = nFilasCuota + 1
                        Application.CutCopyMode = False
                    Else
                        bandera = False
                        temp = j & ":" & j
                        Range(temp).Copy
                        temp = nFilasResto & ":" & nFilasResto
                        wsRestoCuotas.Range(temp).PasteSpecial xlPasteAll
                        nFilasResto = nFilasResto + 1
                        Application.CutCopyMode = False
                    End If
                Else
                    temp = j & ":" & j
                    Range(temp).Copy
                    temp = nFilasResto & ":" & nFilasResto
                    wsRestoCuotas.Range(temp).PasteSpecial xlPasteAll
                    nFilasResto = nFilasResto + 1
                    Application.CutCopyMode = False
                End If
            Next j
            'Reiniciar contadores
            montoCuota = 0
            importeTotal = 0
            dni = Cells(i, 5).Value
            tempPrimero = i
            i = i - 1
        End If
    Next i
    'Trato último
    tempUltimo = nFilas
    montoCuota = importeTotal / 8
    'Cargar Cuotas
    monto = 0
    bandera = True
    For j = tempPrimero To tempUltimo
        If bandera Then
            If Cells(j, 9).Value = 2 Then
                monto = monto - Cells(j, 11).Value
            Else
                monto = monto + Cells(j, 11).Value
            End If
            If monto < montoCuota Then
                temp = j & ":" & j
                Range(temp).Copy
                temp = nFilasCuota & ":" & nFilasCuota
                wsCuota1.Range(temp).PasteSpecial xlPasteAll
                nFilasCuota = nFilasCuota + 1
                Application.CutCopyMode = False
            Else
                bandera = False
                temp = j & ":" & j
                Range(temp).Copy
                temp = nFilasResto & ":" & nFilasResto
                wsRestoCuotas.Range(temp).PasteSpecial xlPasteAll
                nFilasResto = nFilasResto + 1
                Application.CutCopyMode = False
            End If
        Else
            temp = j & ":" & j
            Range(temp).Copy
            temp = nFilasResto & ":" & nFilasResto
            wsRestoCuotas.Range(temp).PasteSpecial xlPasteAll
            nFilasResto = nFilasResto + 1
            Application.CutCopyMode = False
        End If
    Next j
           
            
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    
End Sub


Sub Cuotas()
    Dim wsCuota As Excel.Worksheet
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim nFilasResto As Long
    Dim importeTotal As Double
    Dim montoCuota As Double
    Dim tempPrimero As Integer
    Dim tempUltimo As Integer
    Dim dni As String
    Dim cuota As Integer
    
    
    MsgBox "Debe estar ordenado por DNI.", , "¡Atención!"
    
    'Regresa el control a la hoja de origen
    Sheets(1).Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Fila del encabezado
    nFilasCuota = 2
    cuota = 10
    
    'Tratar al primero
    montoCuota = 0
    importeTotal = 0
    dni = Cells(2, 5).Value
    tempPrimero = 2
    For i = 2 To nFilas
        If dni = Cells(i, 5).Value Then
            'Acumulo
            If Cells(i, 9).Value = 2 Then
                importeTotal = importeTotal - Cells(i, 11).Value
            Else
                importeTotal = importeTotal + Cells(i, 11).Value
            End If
        Else
            'Fin DNI
            tempUltimo = i - 1
            montoCuota = importeTotal / cuota
            numCuota = 1
            'Cargar Cuotas
            monto = 0
            bandera = True
            For j = tempPrimero To tempUltimo
                If Cells(j, 9).Value = 2 Then
                    monto = monto - Cells(j, 11).Value
                Else
                    monto = monto + Cells(j, 11).Value
                End If
                If monto < montoCuota Then
                    Cells(j, nColumnas + 1).Value = numCuota
                Else
                    numCuota = numCuota + 1
                    Cells(j - 1, nColumnas + 2).Value = monto - Cells(j - 1, 11).Value
                    monto = 0
                    j = j - 1
                End If
            Next j
            'Reiniciar contadores
            montoCuota = 0
            importeTotal = 0
            dni = Cells(i, 5).Value
            tempPrimero = i
            i = i - 1
        End If
    Next i
    tempUltimo = i - 1
    montoCuota = importeTotal / cuota
    numCuota = 1
    'Cargar Cuotas
    monto = 0
    bandera = True
    For j = tempPrimero To tempUltimo
        If Cells(j, 9).Value = 2 Then
            monto = monto - Cells(j, 11).Value
        Else
            monto = monto + Cells(j, 11).Value
        End If
        If monto < montoCuota Then
            Cells(j, nColumnas + 1).Value = numCuota
        Else
            numCuota = numCuota + 1
            Cells(j - 1, nColumnas + 2).Value = monto - Cells(j - 1, 11).Value
            monto = 0
            j = j - 1
        End If
    Next j
            
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    
End Sub

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


Sub Cuota_Monto_Max()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Integer
    Dim importeTotal As Double
    Dim montoCuota As Double
    Dim tempPrimero As Integer
    Dim tempUltimo As Integer
    Dim dni As String
    Dim montoMax As Double
    Dim cantCuotas As Integer
    
    montoString = InputBox("Ingrese el importe máximo por couta:", "Jurisdicción", "1")
    montoMax = CDbl(montoString)
    montoCuota = montoMax
    
    MsgBox "Debe estar ordenado por DNI.", , "¡Atención!"
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    Cells(1, nColumnas + 1).Value = "Cuota"
    Cells(1, nColumnas + 2).Value = "Total Cuota"
    
    'Tratar al primero
    importeTotal = 0
    dni = Cells(2, 5).Value
    tempPrimero = 2
    For i = 2 To nFilas
        If dni = Cells(i, 5).Value Then
            'Acumulo
            If Cells(i, 9).Value = 2 Then
                importeTotal = importeTotal - Cells(i, 11).Value
            Else
                importeTotal = importeTotal + Cells(i, 11).Value
            End If
        Else
            'Fin DNI
            tempUltimo = i - 1
            cantCuotas = (importeTotal \ montoMax) + 1
            'Cargar Cuotas
            monto = 0
            cuota = 1
            For j = tempPrimero To tempUltimo
                If Cells(j, 9).Value = 2 Then
                    monto = monto - Cells(j, 11).Value
                Else
                    monto = monto + Cells(j, 11).Value
                End If
                If cuota < cantCuotas Then
                    If monto < montoCuota Then
                        Cells(j, nColumnas + 1) = cuota
                    Else
                        Cells(j - 1, nColumnas + 2) = monto - Cells(j, 11).Value
                        cuota = cuota + 1
                        monto = 0
                        j = j - 1
                    End If
                Else
                    Cells(j, nColumnas + 1) = cuota
                End If
            Next j
            Cells(j - 1, nColumnas + 2) = monto
            
            'Reiniciar contadores
            importeTotal = 0
            dni = Cells(i, 5).Value
            tempPrimero = i
            i = i - 1
        End If
    Next i
    'Trato último
    tempUltimo = i - 1
    cantCuotas = (importeTotal \ montoMax) + 1
    'Cargar Cuotas
    monto = 0
    cuota = 1
    For j = tempPrimero To tempUltimo
        If Cells(j, 9).Value = 2 Then
            monto = monto - Cells(j, 11).Value
        Else
            monto = monto + Cells(j, 11).Value
        End If
        If cuota < cantCuotas Then
            If monto < montoCuota Then
                Cells(j, nColumnas + 1) = cuota
            Else
                Cells(j - 1, nColumnas + 2) = monto - Cells(j, 11).Value
                cuota = cuota + 1
                monto = 0
                j = j - 1
            End If
        Else
            Cells(j, nColumnas + 1) = cuota
        End If
    Next j
    Cells(j - 1, nColumnas + 2) = monto
           
            
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    
End Sub

Sub Cuota_Monto()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Integer
    Dim importeTotal As Double
    Dim montoCuota As Double
    Dim tempPrimero As Integer
    Dim tempUltimo As Integer
    Dim dni As String
    Dim montoMax As Double
    Dim cantCuotas As Integer
    
    montoString = InputBox("Ingrese el importe máximo por couta:", "Jurisdicción", "1")
    montoMax = CDbl(montoString)
    
    MsgBox "Debe estar ordenado por DNI.", , "¡Atención!"
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    Cells(1, nColumnas + 1).Value = "Cuota"
    Cells(1, nColumnas + 2).Value = "Total Cuota"
    
    'Tratar al primero
    montoCuota = 0
    importeTotal = 0
    dni = Cells(2, 5).Value
    tempPrimero = 2
    For i = 2 To nFilas
        If dni = Cells(i, 5).Value Then
            'Acumulo
            If Cells(i, 9).Value = 2 Then
                importeTotal = importeTotal - Cells(i, 11).Value
            Else
                importeTotal = importeTotal + Cells(i, 11).Value
            End If
        Else
            'Fin DNI
            tempUltimo = i - 1
            cantCuotas = (importeTotal \ montoMax) + 1
            montoCuota = importeTotal / cantCuotas
            'Cargar Cuotas
            monto = 0
            cuota = 1
            For j = tempPrimero To tempUltimo
                If Cells(j, 9).Value = 2 Then
                    monto = monto - Cells(j, 11).Value
                Else
                    monto = monto + Cells(j, 11).Value
                End If
                If cuota < cantCuotas Then
                    If monto < montoCuota Then
                        Cells(j, nColumnas + 1) = cuota
                    Else
                        Cells(j - 1, nColumnas + 2) = monto - Cells(j, 11).Value
                        cuota = cuota + 1
                        monto = 0
                        j = j - 1
                    End If
                Else
                    Cells(j, nColumnas + 1) = cuota
                End If
            Next j
            Cells(j - 1, nColumnas + 2) = monto
            
            'Reiniciar contadores
            montoCuota = 0
            importeTotal = 0
            dni = Cells(i, 5).Value
            tempPrimero = i
            i = i - 1
        End If
    Next i
    'Trato último
    tempUltimo = i - 1
    cantCuotas = (importeTotal \ montoMax) + 1
    montoCuota = importeTotal / cantCuotas
    'Cargar Cuotas
    monto = 0
    cuota = 1
    For j = tempPrimero To tempUltimo
        If Cells(j, 9).Value = 2 Then
            monto = monto - Cells(j, 11).Value
        Else
            monto = monto + Cells(j, 11).Value
        End If
        If cuota < cantCuotas Then
            If monto < montoCuota Then
                Cells(j, nColumnas + 1) = cuota
            Else
                Cells(j - 1, nColumnas + 2) = monto - Cells(j, 11).Value
                cuota = cuota + 1
                monto = 0
                j = j - 1
            End If
        Else
            Cells(j, nColumnas + 1) = cuota
        End If
    Next j
    Cells(j - 1, nColumnas + 2) = monto
           
            
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    
End Sub


Sub Crear_Listado()
    Dim wsListado As Excel.Worksheet
    Dim rango As Range
    Dim nFilas As Long
    Dim nFilasLista As Long
    Dim importeTotal As Double
    Dim montoCuota As Double
    Dim tempPrimero As Integer
    Dim tempUltimo As Integer
    Dim dni As String
    
    
    MsgBox "Debe estar ordenado por DNI.", , "¡Atención!"
    
    Application.DisplayAlerts = False
    'Agrega las nuevas hojas
    Worksheets.Add
    ActiveSheet.Name = "Listado"
    Application.DisplayAlerts = True
    
    Set wsListado = Worksheets("Listado")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    'Fila del encabezado
    nFilasLista = 1
    
    wsListado.Cells(nFilasLista, 1).Value = "JUR"
    wsListado.Cells(nFilasLista, 2).Value = "DNI"
    wsListado.Cells(nFilasLista, 3).Value = "NOMBRE"
    wsListado.Cells(nFilasLista, 4).Value = "IMPORTE TOTAL"
    
    nFilasLista = nFilasLista + 1
    
    'Tratar al primero
    importeTotal = 0
    dni = Cells(2, 5).Value
    
    wsListado.Cells(nFilasLista, 1).Value = Cells(2, 2).Value
    wsListado.Cells(nFilasLista, 2).Value = Cells(2, 5).Value
    wsListado.Cells(nFilasLista, 3).Value = Cells(2, 7).Value
    For i = 2 To nFilas
        If dni = Cells(i, 5).Value Then
            'Acumulo
            If Cells(i, 9).Value = 2 Then
                importeTotal = importeTotal - Cells(i, 11).Value
            Else
                importeTotal = importeTotal + Cells(i, 11).Value
            End If
        Else
            'Fin DNI
            wsListado.Cells(nFilasLista, 4).Value = importeTotal
            nFilasLista = nFilasLista + 1
            
            importeTotal = 0
            dni = Cells(i, 5).Value
            wsListado.Cells(nFilasLista, 1).Value = Cells(i, 2).Value
            wsListado.Cells(nFilasLista, 2).Value = Cells(i, 5).Value
            wsListado.Cells(nFilasLista, 3).Value = Cells(i, 7).Value
            i = i - 1
        End If
    Next i
    'Trato último
    wsListado.Cells(nFilasLista, 4).Value = importeTotal
           
            
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    
End Sub

