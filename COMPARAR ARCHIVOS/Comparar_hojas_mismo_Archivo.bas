Attribute VB_Name = "Módulo1"
    Sub Comprar_Hojas_Dif_mismo_Arch()
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
    
    'Calcular el número de filas de la hoja Contenido
     Sheets("walter mes 6").Select
     Set rangoCont = ActiveSheet.UsedRange
     nFilasCont = rangoCont.Rows.Count
     nColumnasCont = rangoCont.Columns.Count
    
    limite = nFilas
    nColumnas = nColumnas + 1
    nColumnasCont = nColumnasCont + 1
    
    For i = 2 To limite
      Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
      
      'Regresa el control a la hoja origen
       Sheets("Hoja1").Select
            
       pos2 = i
       dni = Cells(i, 1).Value
       
       For j = 2 To nFilasCont
       
        'Regresa el control a la hoja nueva
          Sheets("walter mes 6").Select
          
         If Sheets("walter mes 6").Cells(j, 1).Value = dni Then
            Cells(j, nColumnasCont).Value = "encontrado"
            
            'Regresa el control a la hoja origen
             Sheets("Hoja1").Select
                                                     
             Cells(i, nColumnas).Value = "ok en hoja2"
                                             
         End If
                
       Next j
    
    Next i
    
    'Regresa el control a la hoja de origen
     Sheets("Hoja1").Select
    
     MsgBox "Proceso exitoso"
     Application.StatusBar = False

End Sub

   







