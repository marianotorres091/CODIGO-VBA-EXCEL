Attribute VB_Name = "Módulo1"
    Sub Comprar_Hojas_Dif_mismo_Arch()
    Dim rango As Range
    Dim rangoCont As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim nColumnasCont As Long
    Dim i As Long
    Dim filaCopia As Long
    
    
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
  
    'Regresa el control a la hoja de origen
     Sheets("Hoja1").Select
   
    
    'Calcular el número de filas de la hoja actual
     Set rango = ActiveSheet.UsedRange
     nFilas = rango.Rows.Count
     nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
     Sheets("Hoja2").Select
     Set rangoCont = ActiveSheet.UsedRange
     nFilasCont = rangoCont.Rows.Count
     nColumnasCont = rangoCont.Columns.Count
    
    limite = 112
    nColumnas = nColumnas + 1
    nColumnasCont = nColumnasCont + 1
    
    For i = 7 To limite
      Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
      
      'Regresa el control a la hoja origen
       Sheets("Hoja1").Select
            
     
       ceic = Cells(i, 4).Value
      
       For j = 2 To nFilasCont
       
        'Regresa el control a la hoja nueva
          Sheets("Hoja2").Select
          
         If Sheets("Hoja2").Cells(j, 7).Value = ceic Then
           
                    filaResultado = filaResultado + 1
                    Sheets("RESULTADO").Cells(filaResultado, 1).Value = 0
                    Sheets("RESULTADO").Cells(filaResultado, 2).Value = 48
                    Sheets("RESULTADO").Cells(filaResultado, 3).Value = 2
                    Sheets("RESULTADO").Cells(filaResultado, 4).Value = 0
                    Sheets("RESULTADO").Cells(filaResultado, 5).Value = Sheets("Hoja1").Cells(i, 3).Value
                    Sheets("RESULTADO").Cells(filaResultado, 6).Value = 0
                    Sheets("RESULTADO").Cells(filaResultado, 7).Value = Sheets("Hoja1").Cells(i, 2).Value
                    Sheets("RESULTADO").Cells(filaResultado, 8).Value = Sheets("Hoja2").Cells(j, 2).Value
                    Sheets("RESULTADO").Cells(filaResultado, 9).Value = Sheets("Hoja2").Cells(j, 3).Value
                    Sheets("RESULTADO").Cells(filaResultado, 10).Value = Sheets("Hoja2").Cells(j, 4).Value
                    Sheets("RESULTADO").Cells(filaResultado, 11).Value = Sheets("Hoja2").Cells(j, 5).Value
                     Sheets("RESULTADO").Cells(filaResultado, 12).Value = Sheets("Hoja2").Cells(j, 6).Value
            
            
            
            'Regresa el control a la hoja origen
             Sheets("Hoja1").Select
                                                     
             Cells(i, nColumnas).Value = "ok"
                                             
         End If
                
       Next j
    
    Next i
    
    'Regresa el control a la hoja de origen
     Sheets("Hoja1").Select
    
     MsgBox "Proceso exitoso"
     Application.StatusBar = False

End Sub


    Sub TOTALES()
    Dim rango As Range
    Dim rangoCont As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim nColumnasCont As Long
    Dim i As Long
    Dim filaCopia As Long
  
    'Regresa el control a la hoja de origen
     Sheets("Hoja1").Select
     
    
    'Calcular el número de filas de la hoja actual
     Set rango = ActiveSheet.UsedRange
     nFilas = rango.Rows.Count
     nColumnas = rango.Columns.Count
    
     Cells(6, nColumnas + 1).Value = "MES 92019"
     Cells(6, nColumnas + 2).Value = "MES 102019"
     Cells(6, nColumnas + 3).Value = "SAC"
     Cells(6, nColumnas + 4).Value = "TOTAL"
    
    
    
    
    'Calcular el número de filas de la hoja Contenido
     Sheets("Hoja3").Select
     Set rangoCont = ActiveSheet.UsedRange
     nFilasCont = rangoCont.Rows.Count
     nColumnasCont = rangoCont.Columns.Count
    
    limite = nFilas
    nColumnas = nColumnas + 1
    nColumnasCont = nColumnasCont + 1
    
    For i = 7 To 112
      Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
      
      'Regresa el control a la hoja origen
       Sheets("Hoja1").Select
            
       
       ceic = Cells(i, 4).Value
       
       For j = 2 To nFilasCont
       
        'Regresa el control a la hoja nueva
          Sheets("Hoja3").Select
          
         If Sheets("Hoja3").Cells(j, 1).Value = ceic Then
           
            
            'Regresa el control a la hoja origen
             Sheets("Hoja1").Select
                                                     
             Cells(i, nColumnas).Value = Sheets("Hoja3").Cells(j, 2).Value
             Cells(i, nColumnas + 1).Value = Sheets("Hoja3").Cells(j, 3).Value
             Cells(i, nColumnas + 2).Value = Sheets("Hoja3").Cells(j, 4).Value
             Cells(i, nColumnas + 3).Value = Sheets("Hoja3").Cells(j, 5).Value
                                             
         End If
                
       Next j
    
    Next i
    
    'Regresa el control a la hoja de origen
     Sheets("Hoja1").Select
    
     MsgBox "Proceso exitoso"
     Application.StatusBar = False

End Sub

   







