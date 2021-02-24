Attribute VB_Name = "Módulo11"
Sub Obtener_Carga()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim tempFecha As Date
    Dim i, RowNumber As Long
    Dim filaResultado As Long
    Dim filaCopia As Long
    Dim columnaCopia As Long
    Dim wbContenido As Worksheet
       

    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    

     Sheets.Add.Name = "RESULTADO"
 
    
    Cells(1, 1).Value = "PtaId"
    Cells(1, 2).Value = "JurId"
    Cells(1, 3).Value = "EscId"
    Cells(1, 4).Value = "Pref"
    Cells(1, 5).Value = "Doc"
    Cells(1, 6).Value = "Digito"
    Cells(1, 7).Value = "Nombres"
    Cells(1, 8).Value = "Couc"
    Cells(1, 9).Value = "Reajuste"
    Cells(1, 10).Value = "Unidades"
    Cells(1, 11).Value = "Importe"
    Cells(1, 12).Value = "Vto"
        
   filaResultado = 1
    
    
    For i = 5 To 94
       
                
                
                If Sheets("Hoja1").Cells(i, 1) = "" Then
                    filaResultado = filaResultado + 1
                    Sheets("RESULTADO").Cells(filaResultado, 1).Value = 0
                    Sheets("RESULTADO").Cells(filaResultado, 2).Value = Sheets("Hoja1").Cells(i, 3).Value
                    Sheets("RESULTADO").Cells(filaResultado, 3).Value = 2
                    Sheets("RESULTADO").Cells(filaResultado, 4).Value = 0
                    Sheets("RESULTADO").Cells(filaResultado, 5).Value = Sheets("Hoja1").Cells(i, 5).Value
                    Sheets("RESULTADO").Cells(filaResultado, 6).Value = 0
                    Sheets("RESULTADO").Cells(filaResultado, 7).Value = Sheets("Hoja1").Cells(i, 6).Value
                    Sheets("RESULTADO").Cells(filaResultado, 8).Value = Sheets("Hoja1").Cells(i, 7).Value
                    Sheets("RESULTADO").Cells(filaResultado, 9).Value = 1
                    Sheets("RESULTADO").Cells(filaResultado, 10).Value = Sheets("Hoja1").Cells(i, 9).Value
                    Sheets("RESULTADO").Cells(filaResultado, 11).Value = Sheets("Hoja1").Cells(i, 10).Value
                     Sheets("RESULTADO").Cells(filaResultado, 12).Value = Sheets("Hoja1").Cells(i, 11).Value
               End If
    Next i
    
                
       
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub

End Sub



Sub contar()
 RowNumber = ActiveSheet.Range("L132000").End(xlUp).Row
 MsgBox (RowNumber)
End Sub

Sub TOTAL()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i, TOTAL As Long
    

    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    For i = 2 To nFilas
      Sheets("RESULTADO").Cells(nFilas, 13).Value = Sheets("RESULTADO").Cells(nFilas, 13).Value + Sheets("RESULTADO").Cells(i, 11).Value
    Next i
   
    MsgBox "Proceso Exitoso el total es", , Cells(nFilas, 13).Value
End Sub


