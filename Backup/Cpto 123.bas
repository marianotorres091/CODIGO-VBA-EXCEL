Attribute VB_Name = "Module1"
Sub Controlar_Pagados()
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
    Dim tempMes As Integer
    Dim tempAnio As Integer
    
    
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
    
    wsResultado.Cells(1, 1).Value = "Jur"
    wsResultado.Cells(1, 2).Value = "DNI"
    wsResultado.Cells(1, 3).Value = "Nombre"
    wsResultado.Cells(1, 4).Value = "Cant Meses Pagos"
    wsResultado.Cells(1, 5).Value = "Cant Meses Desc"
    wsResultado.Cells(1, 6).Value = "Importe Total"
    wsResultado.Cells(1, 7).Value = "Último Mes"
    nFilasCont = 2
    
    valorDoc = Cells(2, 12).Value
    importeTotal = Cells(2, 7).Value
    contador = 1
    contador2 = 0
    
    wsResultado.Cells(nFilasCont, 1).Value = Cells(2, 8).Value
    wsResultado.Cells(nFilasCont, 2).Value = valorDoc
    wsResultado.Cells(nFilasCont, 3).Value = Cells(2, 14).Value
    
    For i = 3 To nFilas
        If valorDoc = Cells(i, 12).Value Then
            If Cells(i, 6).Value = 0 Then
                tempMes = Cells(i, 2).Value
                tempAnio = Cells(i, 1).Value
            Else
                valorVto = Cells(i, 16).Value
                For m = 1 To Len(valorVto)
                    If Mid(valorVto, m, 1) = "/" Then
                        tempMes = Mid(valorVto, m + 1, 1)
                        m = m + 2
                        If Mid(valorVto, m, 1) <> "/" Then
                            tempMes = tempMes & Mid(valorVto, m, 1)
                            m = m + 1
                        End If
                        tempAnio = Mid(valorVto, m + 1, 4)
                        m = 10
                    End If
                Next m
                If tempAnio > Cells(i, 1).Value Then
                    tempAnio = Cells(i, 1).Value
                    tempMes = Cells(i, 2).Value
                Else
                    If tempMes > Cells(i, 2).Value Then
                        tempMes = Cells(i, 2).Value
                    End If
                End If
            End If
            
            If Cells(i, 6).Value = 2 Then
                importeTotal = importeTotal - Cells(i, 7).Value
                contador2 = contador2 + 1
            Else
                importeTotal = importeTotal + Cells(i, 7).Value
                contador = contador + 1
            End If

        Else
            'Copio en archivo
            wsResultado.Cells(nFilasCont, 4).Value = contador
            wsResultado.Cells(nFilasCont, 5).Value = contador2
            wsResultado.Cells(nFilasCont, 6).Value = importeTotal
            wsResultado.Cells(nFilasCont, 7).Value = tempMes & " - " & tempAnio
            nFilasCont = nFilasCont + 1
            
            'Reinicio contadores
            valorDoc = Cells(i, 12).Value
            contador = 0
            contador2 = 0
            importeTotal = 0
            wsResultado.Cells(nFilasCont, 1).Value = Cells(i, 8).Value
            wsResultado.Cells(nFilasCont, 2).Value = valorDoc
            wsResultado.Cells(nFilasCont, 3).Value = Cells(i, 14).Value
            i = i - 1
        End If
    Next i
    
    wsResultado.Cells(nFilasCont, 4).Value = contador
    wsResultado.Cells(nFilasCont, 5).Value = contador2
    wsResultado.Cells(nFilasCont, 6).Value = importeTotal
    wsResultado.Cells(nFilasCont, 7).Value = tempMes & " - " & tempAnio
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub

Sub Listar_UltVto()
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
    Dim tempMes As Integer
    Dim tempAnio As Integer
    Dim hoy As Date
    
    
    'Activar este libro
    ThisWorkbook.Activate

    'Agrega las nuevas hojas
    Worksheets.Add
    ActiveSheet.Name = "UltVto"
    Application.DisplayAlerts = True
    
    Set wsResultado = Worksheets("UltVto")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select

    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    nFilasCont = 1
    wsResultado.Cells(nFilasCont, 1).Value = "PtaId"
    wsResultado.Cells(nFilasCont, 2).Value = "JurId"
    wsResultado.Cells(nFilasCont, 3).Value = "EscId"
    wsResultado.Cells(nFilasCont, 4).Value = "Pref"
    wsResultado.Cells(nFilasCont, 5).Value = "Doc"
    wsResultado.Cells(nFilasCont, 6).Value = "Digito"
    wsResultado.Cells(nFilasCont, 7).Value = "Nombres"
    wsResultado.Cells(nFilasCont, 8).Value = "Couc"
    wsResultado.Cells(nFilasCont, 9).Value = "Reajuste"
    wsResultado.Cells(nFilasCont, 10).Value = "Unidades"
    wsResultado.Cells(nFilasCont, 11).Value = "Importe"
    wsResultado.Cells(nFilasCont, 12).Value = "Vto"
    
    valorDoc = "0"
    hoy = DateValue("May 17, 2018")
    
    For i = 3 To nFilas
        If Cells(i, 16).Value > hoy Then
            If valorDoc <> Cells(i, 12).Value Then
                valorDoc = Cells(i, 12).Value
                nFilasCont = nFilasCont + 1
                wsResultado.Cells(nFilasCont, 1).Value = 0
                wsResultado.Cells(nFilasCont, 2).Value = Cells(i, 8).Value
                wsResultado.Cells(nFilasCont, 3).Value = Cells(i, 9).Value
                wsResultado.Cells(nFilasCont, 4).Value = 0
                wsResultado.Cells(nFilasCont, 5).Value = valorDoc
                wsResultado.Cells(nFilasCont, 6).Value = 0
                wsResultado.Cells(nFilasCont, 7).Value = Cells(i, 14).Value
                wsResultado.Cells(nFilasCont, 8).Value = 123
                wsResultado.Cells(nFilasCont, 9).Value = 1
                wsResultado.Cells(nFilasCont, 10).Value = 0
                wsResultado.Cells(nFilasCont, 11).Value = Cells(i, 7).Value
                
                valorVto = Cells(i, 16).Value
                For m = 1 To Len(valorVto)
                    If Mid(valorVto, m, 1) = "/" Then
                        tempMes = Mid(valorVto, m + 1, 1)
                        m = m + 2
                        If Mid(valorVto, m, 1) <> "/" Then
                            tempMes = tempMes & Mid(valorVto, m, 1)
                            m = m + 1
                        End If
                        tempAnio = Mid(valorVto, m + 1, 4)
                        m = 10
                    End If
                Next m
                
                wsResultado.Cells(nFilasCont, 12).Value = tempMes & tempAnio
            End If
        End If
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub


