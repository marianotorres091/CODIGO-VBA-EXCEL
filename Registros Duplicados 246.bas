Attribute VB_Name = "Módulo1"
Sub Registros_duplicados_246()
    Dim nFilas As Long
    Dim nColumnas As Long
    'Dim i As Integer
    Dim rango As Range
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
   
    limite = (nFilas - 1)
    For i = 2 To limite
     Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
     If Cells(i, nColumnas + 1).Value <> "Repetido" Then
        valorjur = Cells(i, 1).Value
        valorDoc = Cells(i, 4).Value
        valormes = Cells(i, 8).Value
        valortipo = Cells(i, 9).Value
        valornro = Cells(i, 10).Value
        valorbase = Cells(i, 11).Value
        valorReajCpto_025 = Cells(i, 12).Value
        valorImpCpto_025 = Cells(i, 13).Value
        valorReajCpto_246 = Cells(i, 14).Value
        valorUnidCpto_246 = Cells(i, 15).Value
        valorImpCpto_246 = Cells(i, 16).Value
        valorPorcRet = Cells(i, 17).Value
        valorDiasTrab = Cells(i, 18).Value
        valorDiasComplemento = Cells(i, 19).Value
        valorimporte = Cells(i, 20).Value
        valorDiferencia = Cells(i, 21).Value
        valorCpto = Cells(i, 22).Value
        
       
        For j = (i + 1) To nFilas
            If valorjur = Cells(j, 1).Value And valorDoc = Cells(j, 4).Value And valormes = Cells(j, 8).Value And valortipo = Cells(j, 9).Value And valornro = Cells(j, 10).Value And valorbase = Cells(j, 11).Value And valorReajCpto_025 = Cells(j, 12).Value And valorImpCpto_025 = Cells(j, 13).Value And valorReajCpto_246 = Cells(j, 14).Value And valorUnidCpto_246 = Cells(j, 15).Value And valorImpCpto_246 = Cells(j, 16).Value And valorPorcRet = Cells(j, 17).Value And valorDiasTrab = Cells(j, 18).Value And valorDiasComplemento = Cells(j, 19).Value And valorimporte = Cells(j, 20).Value And valorDiferencia = Cells(j, 21).Value And valorCpto = Cells(j, 22).Value Then
                Cells(j, 4).Interior.Color = RGB(153, 196, 195)
                Cells(i, 4).Interior.Color = RGB(51, 255, 90)
                Cells(j, nColumnas + 1).Value = "Repetido"
                Cells(j, nColumnas + 2).Value = i
                Cells(i, nColumnas + 2).Value = i
                Cells(i, nColumnas + 1).Value = "Repetido"
            End If
        Next j
      End If
    Next i
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Application.StatusBar = False
End Sub








