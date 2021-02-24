Attribute VB_Name = "Módulo11"
Sub BuscarEnVariasHojas()
Dim Celda As Range, rango As Range
Dim CeldaAux As Range, RangoAux As Range
Dim Hoja As Worksheet



Hoja1.Activate

Set rango = Hoja1.Range(Cells(2, 4), Cells(2, 4).End(xlDown))

For Each Hoja In ActiveWorkbook.Worksheets
    If Hoja.Name <> "HISTORICO" Then
        Hoja.Activate
        
        For Each Celda In rango
            Set RangoAux = Hoja.Range(Cells(2, 4), Cells(2, 4).End(xlDown))
            For Each CeldaAux In RangoAux
                If Celda = CeldaAux Then
                    Celda.Offset(0, 18) = "si"
                   
                End If
            Next CeldaAux

        Next Celda
    End If
Next Hoja

Hoja1.Activate


MsgBox "Proceso Exitoso "

End Sub

Sub copiar()
Dim i, j, x, y As Integer
Dim nFilas As Long
Dim nColumnas As Long
  
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count


x = 1
y = 12
For i = 2 To nFilas
 For j = 1 To 8
   If Cells(i, 10) = "si" Then
   Cells(x, y).Value = Cells(i, j).Value
    y = y + 1
    If j = 8 Then
     x = x + 1
     y = 12
    End If
   End If
 Next j

 
Next i

MsgBox "Proceso Exitoso "
End Sub


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
    
    
    For i = 2 To nFilas
       
                
                
                If Sheets("HISTORICO").Cells(i, 22) = "si" Then
                    filaResultado = filaResultado + 1
                    Sheets("RESULTADO").Cells(filaResultado, 1).Value = 0
                    Sheets("RESULTADO").Cells(filaResultado, 2).Value = Sheets("HISTORICO").Cells(i, 2).Value
                    Sheets("RESULTADO").Cells(filaResultado, 3).Value = 2
                    Sheets("RESULTADO").Cells(filaResultado, 4).Value = 0
                    Sheets("RESULTADO").Cells(filaResultado, 5).Value = Sheets("HISTORICO").Cells(i, 4).Value
                    Sheets("RESULTADO").Cells(filaResultado, 6).Value = 0
                    Sheets("RESULTADO").Cells(filaResultado, 7).Value = Sheets("HISTORICO").Cells(i, 6).Value
                    Sheets("RESULTADO").Cells(filaResultado, 8).Value = 246
                    Sheets("RESULTADO").Cells(filaResultado, 9).Value = 1
                    Sheets("RESULTADO").Cells(filaResultado, 10).Value = 0
                    Sheets("RESULTADO").Cells(filaResultado, 11).Value = Sheets("HISTORICO").Cells(i, 21).Value
                     Sheets("RESULTADO").Cells(filaResultado, 12).Value = "92019"
               End If
    Next i
    
                
       
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Exit Sub

End Sub



Sub contar()
 RowNumber = ActiveSheet.Range("L132000").End(xlUp).Row
 MsgBox (RowNumber)
End Sub



