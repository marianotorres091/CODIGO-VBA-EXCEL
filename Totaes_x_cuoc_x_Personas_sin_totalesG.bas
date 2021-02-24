Attribute VB_Name = "Módulo2"
Sub Totales_x_cuoc__x_Persona2()
    Dim nFilas As Long
    Dim filaTotal As Long
    Dim rango As Range
    Dim wsTotal As Excel.Worksheet
    Dim i As Long
    Dim ultDoc As String
    Dim ultJur As Integer
    Dim band As Double
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Total x cuoc x Persona"
    Application.DisplayAlerts = True
    
    Set wsTotal = Worksheets("Total x cuoc x Persona")
    
      
     'Regresa el control a la hoja de origen
    Sheets("HISTORICO").Select
    
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    

    
    'MsgBox "Debe estar ordenado por DNI.", , "Atención!!"
    
    'Encabezado Hoja Totales
    wsTotal.Cells(1, 1).Value = "JUR"
    wsTotal.Cells(1, 2).Value = "DNI"
    wsTotal.Cells(1, 3).Value = "Nombre"
    wsTotal.Cells(1, 4).Value = "Cuoc"
    wsTotal.Cells(1, 5).Value = "Cuoc-Reaj 1"
    wsTotal.Cells(1, 6).Value = "Cuoc-Reaj 2"
    filaTotal = 2
     
    vto = Cells(2, 12).Value
    ultDoc = Cells(2, 5).Value
    importe = 0
    ultJur = Cells(2, 2).Value
    nombre = Cells(2, 7).Value
    band = False
    pos = 2
    totalimportemas = 0
    totalimportemenos = 0
    
    For i = 2 To nFilas
    
    Sheets("HISTORICO").Select
    
      If ultDoc = Cells(i, 5).Value Then
        cuoc = Cells(i, 8).Value
       'Trato el primero
       If i = 2 Then
             Sheets("Total x cuoc x Persona").Select
             wsTotal.Cells(filaTotal, 1).Value = ultJur
             wsTotal.Cells(filaTotal, 2).Value = ultDoc
             wsTotal.Cells(filaTotal, 3).Value = nombre
             wsTotal.Cells(filaTotal, 4).Value = cuoc
             wsTotal.Cells(filaTotal, 7).Value = vto
             Sheets("HISTORICO").Select
            'segun el reajuste sumo en la hoja Total_x_cuoc_Persona del primer registro
            If Cells(i, 9).Value = 1 Then
              Sheets("Total x cuoc x Persona").Select
              wsTotal.Cells(filaTotal, 5).Value = wsTotal.Cells(filaTotal, 5).Value + Sheets("HISTORICO").Cells(i, 11).Value
              totalimportemas = totalimportemas + Sheets("HISTORICO").Cells(i, 11).Value
             Else
              Sheets("Total x cuoc x Persona").Select
              wsTotal.Cells(filaTotal, 6).Value = wsTotal.Cells(filaTotal, 6).Value - Sheets("HISTORICO").Cells(i, 11).Value
              totalimportemenos = totalimportemenos - Sheets("HISTORICO").Cells(i, 11).Value
            End If
            
             Sheets("Total x cuoc x Persona").Select
             filaTotal = filaTotal + 1
             Sheets("HISTORICO").Select
         Else
           
            Sheets("Total x cuoc x Persona").Select
          
            'Calculo el número de filas de la hoja de los cobrados
             Set rangoCont = ActiveSheet.UsedRange
             nFilasCont = rangoCont.Rows.Count
             
            'Aca hago la busqueda de si existe ese cuoc en la hoja Total x cuoc x Persona
           For j = pos To nFilasCont
            If Sheets("Total x cuoc x Persona").Cells(j, 4).Value <> cuoc Then
              
             Else
              band = True
              pos2 = j
            End If
            Next j
        
          If band = False Then
          
                  Sheets("HISTORICO").Select
                  'segun el rajuste lo acumulo
                If Cells(i, 9).Value = 1 Then
                  Sheets("Total x cuoc x Persona").Select
                  wsTotal.Cells(filaTotal, 5).Value = wsTotal.Cells(filaTotal, 5).Value + Sheets("HISTORICO").Cells(i, 11).Value
                  totalimportemas = totalimportemas + Sheets("HISTORICO").Cells(i, 11).Value
                 Else
                  Sheets("Total x cuoc x Persona").Select
                  wsTotal.Cells(filaTotal, 6).Value = wsTotal.Cells(filaTotal, 6).Value - Sheets("HISTORICO").Cells(i, 11).Value
                totalimportemenos = totalimportemenos - Sheets("HISTORICO").Cells(i, 11).Value
                End If
                  'almaceno el cuoc en la hoja Total x cuoc x Persona
                 
                  wsTotal.Cells(filaTotal, 4).Value = cuoc
                  wsTotal.Cells(filaTotal, 7).Value = vto
                  cont = cont + 1
                  filaTotal = filaTotal + 1
            Else
                  Sheets("HISTORICO").Select
                  'segun el rajuste lo acumulo
                If Cells(i, 9).Value = 1 Then
                  Sheets("Total x cuoc x Persona").Select
                  wsTotal.Cells(pos2, 5).Value = wsTotal.Cells(pos2, 5).Value + Sheets("HISTORICO").Cells(i, 11).Value
                  totalimportemas = totalimportemas + Sheets("HISTORICO").Cells(i, 11).Value
                 Else
                  Sheets("Total x cuoc x Persona").Select
                  wsTotal.Cells(pos2, 6).Value = wsTotal.Cells(pos2, 6).Value - Sheets("HISTORICO").Cells(i, 11).Value
                  totalimportemenos = totalimportemenos - Sheets("HISTORICO").Cells(i, 11).Value
                End If
            
              band = False
             
          End If
          
        End If
        
       Else
              'almaceno el ultimo
              pos = filaTotal
              
              Sheets("Total x cuoc x Persona").Select
              wsTotal.Cells(filaTotal, 1).Value = Sheets("HISTORICO").Cells(i, 2).Value
              wsTotal.Cells(filaTotal, 2).Value = Sheets("HISTORICO").Cells(i, 5).Value
              wsTotal.Cells(filaTotal, 3).Value = Sheets("HISTORICO").Cells(i, 7).Value
              wsTotal.Cells(filaTotal, 4).Value = Sheets("HISTORICO").Cells(i, 8).Value
              wsTotal.Cells(filaTotal, 7).Value = Sheets("HISTORICO").Cells(i, 12).Value
             
              
              Sheets("HISTORICO").Select
              
            If Cells(i, 9).Value = 1 Then
              Sheets("Total x cuoc x Persona").Select
              wsTotal.Cells(filaTotal, 5).Value = wsTotal.Cells(filaTotal, 5).Value + Sheets("HISTORICO").Cells(i, 11).Value
              totalimportemas = totalimportemas + Sheets("HISTORICO").Cells(i, 11).Value
             Else
              Sheets("Total x cuoc x Persona").Select
              wsTotal.Cells(filaTotal, 6).Value = wsTotal.Cells(filaTotal, 6).Value - Sheets("HISTORICO").Cells(i, 11).Value
              totalimportemenos = totalimportemenos - Sheets("HISTORICO").Cells(i, 11).Value
             End If
             filaTotal = filaTotal + 1
              Sheets("Total x cuoc x Persona").Select
              
            
              'asigno el nuevo registro
              Sheets("HISTORICO").Select
              ultDoc = Sheets("HISTORICO").Cells(i, 5).Value
              ultJur = Sheets("HISTORICO").Cells(i, 2).Value
              nombre = Sheets("HISTORICO").Cells(i, 7).Value
              cuoc = Sheets("HISTORICO").Cells(i, 8).Value
              vto = Sheets("HISTORICO").Cells(i, 12).Value
      End If
    Next i
 
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub










