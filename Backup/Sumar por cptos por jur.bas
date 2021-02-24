Attribute VB_Name = "Módulo1"
Sub Totales_x_cpto_x_Jur()
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
    ActiveSheet.Name = "TOTAL_CPTO_jUR"
    Application.DisplayAlerts = True
    
    Set wsTotal = Worksheets("TOTAL_CPTO_jUR")
    
      
     'Regresa el control a la hoja de origen
    Sheets("HISTORICO").Select
    
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    

    
    'MsgBox "Debe estar ordenado por DNI.", , "Atención!!"
    
    'Encabezado Hoja Totales
    wsTotal.Cells(1, 1).Value = "JUR"
    wsTotal.Cells(1, 2).Value = "CPTO"
    wsTotal.Cells(1, 3).Value = "DESCRIPCION"
    wsTotal.Cells(1, 4).Value = "IMPORTE"
    
    filaTotal = 2
     
    
    importe = 0
    ultJur = Cells(2, 3).Value
    cuoc = Cells(2, 9).Value
    band = False
    pos = 2
    totalimporte = 0
   
    
    limite = nFilas
    
    For i = 2 To limite
    Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
    Sheets("HISTORICO").Select
    
      If ultJur = Cells(i, 3).Value Then
        cuoc = Cells(i, 9).Value
       
       'Trato el primero
       If i = 2 Then
             Sheets("TOTAL_CPTO_jUR").Select
             wsTotal.Cells(filaTotal, 1).Value = ultJur
             wsTotal.Cells(filaTotal, 2).Value = cuoc
           
             
             Sheets("HISTORICO").Select
            'segun el reajuste sumo en la hoja Total_x_cuoc_Persona del primer registro
            If Cells(i, 10).Value = 1 Then
              Sheets("TOTAL_CPTO_jUR").Select
              wsTotal.Cells(filaTotal, 4).Value = wsTotal.Cells(filaTotal, 4).Value + Sheets("HISTORICO").Cells(i, 12).Value
              totalimporte = totalimporte + Sheets("HISTORICO").Cells(i, 12).Value
             Else
              Sheets("TOTAL_CPTO_jUR").Select
              wsTotal.Cells(filaTotal, 4).Value = wsTotal.Cells(filaTotal, 4).Value - Sheets("HISTORICO").Cells(i, 12).Value
              totalimporte = totalimporte - Sheets("HISTORICO").Cells(i, 12).Value
            End If
            
             Sheets("TOTAL_CPTO_jUR").Select
             filaTotal = filaTotal + 1
             Sheets("HISTORICO").Select
         Else
           
            Sheets("TOTAL_CPTO_jUR").Select
          
            'Calculo el número de filas de la hoja de los cobrados
             Set rangoCont = ActiveSheet.UsedRange
             nFilasCont = rangoCont.Rows.Count
             
            'Aca hago la busqueda de si existe ese cuoc en la hoja TOTAL_CPTO_jUR
           For j = pos To nFilasCont
            If Sheets("TOTAL_CPTO_jUR").Cells(j, 2).Value <> cuoc Then
              
             Else
              band = True
              pos2 = j
            End If
            Next j
        
          If band = False Then
          
                  Sheets("HISTORICO").Select
                  'segun el rajuste lo acumulo
                If Cells(i, 10).Value = 1 Then
                  Sheets("TOTAL_CPTO_jUR").Select
                  wsTotal.Cells(filaTotal, 4).Value = wsTotal.Cells(filaTotal, 4).Value + Sheets("HISTORICO").Cells(i, 12).Value
                  totalimporte = totalimporte + Sheets("HISTORICO").Cells(i, 12).Value
                 Else
                  Sheets("TOTAL_CPTO_jUR").Select
                  wsTotal.Cells(filaTotal, 4).Value = wsTotal.Cells(filaTotal, 4).Value - Sheets("HISTORICO").Cells(i, 12).Value
                totalimporte = totalimporte - Sheets("HISTORICO").Cells(i, 12).Value
                End If
                  'almaceno el cuoc en la hoja TOTAL_CPTO_jUR
                  
                  wsTotal.Cells(filaTotal, 2).Value = cuoc
                  
                  cont = cont + 1
                  filaTotal = filaTotal + 1
            Else
                  Sheets("HISTORICO").Select
                  'segun el rajuste lo acumulo
                If Cells(i, 10).Value = 1 Then
                  Sheets("TOTAL_CPTO_jUR").Select
                  wsTotal.Cells(pos2, 4).Value = wsTotal.Cells(pos2, 4).Value + Sheets("HISTORICO").Cells(i, 12).Value
                  totalimporte = totalimporte + Sheets("HISTORICO").Cells(i, 12).Value
                 Else
                  Sheets("TOTAL_CPTO_jUR").Select
                  wsTotal.Cells(pos2, 4).Value = wsTotal.Cells(pos2, 4).Value - Sheets("HISTORICO").Cells(i, 12).Value
                  totalimporte = totalimporte - Sheets("HISTORICO").Cells(i, 12).Value
                End If
            
              band = False
             
          End If
          
        End If
        
       Else
           'almaceno el ultimo
            pos = filaTotal
            wsTotal.Cells(filaTotal, 3).Value = "TOTAL"
            wsTotal.Cells(filaTotal, 4).Value = totalimporte
            filaTotal = filaTotal + 1
           
           totalimporte = 0
           ultJur = Sheets("HISTORICO").Cells(i, 3).Value
           cuoc = Sheets("HISTORICO").Cells(i, 9).Value
           
           Sheets("TOTAL_CPTO_jUR").Select
          
            'Calculo el número de filas de la hoja de los cobrados
             Set rangoCont = ActiveSheet.UsedRange
             nFilasCont = rangoCont.Rows.Count
             
            'Aca hago la busqueda de si existe ese cuoc en la hoja TOTAL_CPTO_jUR
           For j = pos To nFilasCont
            If Sheets("TOTAL_CPTO_jUR").Cells(j, 2).Value <> cuoc Then
              
             Else
              band = True
              pos2 = j
            End If
            Next j
        
          If band = False Then
          
                  Sheets("HISTORICO").Select
                  'segun el rajuste lo acumulo
                If Cells(i, 10).Value = 1 Then
                  Sheets("TOTAL_CPTO_jUR").Select
                  wsTotal.Cells(filaTotal, 4).Value = wsTotal.Cells(filaTotal, 4).Value + Sheets("HISTORICO").Cells(i, 12).Value
                  totalimporte = totalimporte + Sheets("HISTORICO").Cells(i, 12).Value
                 Else
                  Sheets("TOTAL_CPTO_jUR").Select
                  wsTotal.Cells(filaTotal, 4).Value = wsTotal.Cells(filaTotal, 4).Value - Sheets("HISTORICO").Cells(i, 12).Value
                totalimporte = totalimporte - Sheets("HISTORICO").Cells(i, 12).Value
                End If
                  'almaceno el cuoc en la hoja TOTAL_CPTO_jUR
                 wsTotal.Cells(filaTotal, 1).Value = ultJur
                  wsTotal.Cells(filaTotal, 2).Value = cuoc
                  
                  cont = cont + 1
                  filaTotal = filaTotal + 1
            Else
                  Sheets("HISTORICO").Select
                  'segun el rajuste lo acumulo
                If Cells(i, 10).Value = 1 Then
                  Sheets("TOTAL_CPTO_jUR").Select
                  wsTotal.Cells(pos2, 4).Value = wsTotal.Cells(pos2, 4).Value + Sheets("HISTORICO").Cells(i, 12).Value
                  totalimporte = totalimporte + Sheets("HISTORICO").Cells(i, 12).Value
                 Else
                  Sheets("TOTAL_CPTO_jUR").Select
                  wsTotal.Cells(pos2, 4).Value = wsTotal.Cells(pos2, 4).Value - Sheets("HISTORICO").Cells(i, 12).Value
                  totalimporte = totalimporte - Sheets("HISTORICO").Cells(i, 12).Value
                End If
            
              band = False
             
          End If
          
       
              'almaceno el ultimo
              pos = filaTotal
              wsTotal.Cells(filaTotal, 3).Value = "TOTAL"
              wsTotal.Cells(filaTotal, 4).Value = totalimporte
           
              'Sheets("TOTAL_CPTO_jUR").Select
              'wsTotal.Cells(filaTotal + 1, 1).Value = Sheets("HISTORICO").Cells(i, 3).Value
              'wsTotal.Cells(filaTotal + 1, 2).Value = Sheets("HISTORICO").Cells(i, 9).Value
              'wsTotal.Cells(filaTotal + 1, 3).Value = Sheets("HISTORICO").Cells(i, 7).Value
              'wsTotal.Cells(filaTotal + 1, 4).Value = Sheets("HISTORICO").Cells(i, 8).Value
              'wsTotal.Cells(filaTotal + 1, 7).Value = Sheets("HISTORICO").Cells(i, 12).Value
              totalimporte = 0
            
              'Sheets("HISTORICO").Select
             'filaTotal = filaTotal + 1
            'If Cells(i, 10).Value = 1 Then
              'Sheets("TOTAL_CPTO_jUR").Select
              'wsTotal.Cells(filaTotal, 4).Value = wsTotal.Cells(filaTotal, 4).Value + Sheets("HISTORICO").Cells(i, 12).Value
              'totalimporte = totalimporte + Sheets("HISTORICO").Cells(i, 11).Value
             'Else
              'Sheets("TOTAL_CPTO_jUR").Select
              'wsTotal.Cells(filaTotal, 4).Value = wsTotal.Cells(filaTotal, 4).Value - Sheets("HISTORICO").Cells(i, 12).Value
              'totalimporte = totalimporte - Sheets("HISTORICO").Cells(i, 12).Value
             'End If
             
             filaTotal = filaTotal + 1
             Sheets("TOTAL_CPTO_jUR").Select
              
            
              'asigno el nuevo registro
              'Sheets("HISTORICO").Select

              'ultJur = Sheet   s("HISTORICO").Cells(i, 3).Value
              'cuoc = Sheets("HISTORICO").Cells(i, 9).Value
              
      End If
    Next i
   
              'wsTotal.Cells(filaTotal, 3).Value = "TOTAL"
              'wsTotal.Cells(filaTotal, 4).Value = totalimporte
              
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Application.StatusBar = False
End Sub

Sub prueba()
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
    
    'Application.DisplayAlerts = False
    'Worksheets.Add
    'ActiveSheet.Name = "Total x cuoc x Persona"
    'Application.DisplayAlerts = True
    
    'Set wsTotal = Worksheets("Total x cuoc x Persona")
    
      
     'Regresa el control a la hoja de origen
    Sheets("Hoja2").Select
    
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    x = 1
    Sheets("Hoja2").Range("A" & x + 1 & ":O" & x + 1).Insert
End Sub

