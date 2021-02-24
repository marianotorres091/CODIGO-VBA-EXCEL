Attribute VB_Name = "Módulo1"
Sub Registros_duplicados_dni_cpto_rj_unid_imp_vto_difjur()
    Dim nFilas As Long
    Dim nColumnas As Long
    'Dim i As Integer
    Dim rango As Range
    
    'Busco los cptos que se hayan liquidado para el mismo agente en diferentes jur
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
   
    limite = (nFilas - 1)
    For i = 2 To limite
     Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
     If Cells(i, 20).Value <> "Duplicado" And Cells(i, 20).Value <> "misma jur" Then
        valorDoc = Cells(i, 9).Value
        valorVto = Cells(i, 16).Value
        valorCouc = Cells(i, 12).Value
        valorjur = Cells(i, 6).Value
        valorRj = Cells(i, 13).Value
        valorunidad = Cells(i, 14).Value
        valorimporte = Cells(i, 15).Value
        'valorescalafon = Cells(I, 7).Value
        For j = (i + 1) To nFilas
            If Cells(j, 9).Value = valorDoc And valorVto = Cells(j, 16).Value And valorRj = Cells(j, 13).Value And valorunidad = Cells(j, 14).Value And valorimporte = Cells(j, 15).Value And valorCouc = Cells(j, 12).Value Then
               If valorjur <> Cells(j, 6).Value Then
                Cells(j, 9).Interior.Color = RGB(153, 196, 195)
                Cells(i, 9).Interior.Color = RGB(153, 196, 195)
                Cells(j, 20).Value = "Duplicado"
                Cells(j, 21).Value = i
                Cells(i, 21).Value = i
                Cells(i, 20).Value = "Duplicado"
                Else
                 Cells(j, 20).Value = "misma jur"
               End If
            End If
        Next j
      End If
    Next i
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Application.StatusBar = False
End Sub


