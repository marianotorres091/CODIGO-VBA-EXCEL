Attribute VB_Name = "Módulo1"
Sub Cuota_Monto_Maximo()
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
    
    limite = nFilas
    
    For i = 2 To limite
     Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
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
                     If Cells(j - 1, nColumnas + 2) = "" Then
                        Cells(j - 1, nColumnas + 2) = monto - Cells(j, 11).Value
                     End If
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
              If Cells(j - 1, nColumnas + 2) = "" Then
                Cells(j - 1, nColumnas + 2) = monto - Cells(j, 11).Value
              End If
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
    Application.StatusBar = False
End Sub
