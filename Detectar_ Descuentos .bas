Attribute VB_Name = "Módulo1"
Sub Detectar_si_ajustes_Persona_es_todo_DESCUENTO()
    Dim total_mov As Long
    Dim total_dni As Long
    Dim total_actuaciones As Long
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim filaTotal As Long
    Dim rango As Range
    Dim wsTotal As Excel.Worksheet
    Dim i As Long
    Dim ultDoc As String
    Dim ultJur As Integer
 
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    MsgBox "Debe estar ordenado por DNI.", , "Atención!!"
    
    'Inicializamos variables
    act = Cells(2, 14).Value
    ultDoc = Cells(2, 5).Value
    importe = 0
    ultJur = Cells(2, 2).Value
    nombre = Cells(2, 7).Value
    limite = nFilas
    Cells(1, nColumnas + 5).Value = 0
    pos = 2
    
    For i = 2 To limite
    Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
        If Cells(i, 4).Value < 350 Then
            If ultDoc = Cells(i, 5).Value Then
               If act = Cells(i, 14).Value Then
               
                    If Cells(i, 9).Value = 2 Then
                     Cells(i, nColumnas + 4).Value = 0
                     Cells(1, nColumnas + 5).Value = Cells(1, nColumnas + 5).Value + Cells(i, nColumnas + 4).Value
                    Else
                        band = True
                       Cells(i, nColumnas + 2).Value = "ajuste en mas"
                       Cells(i, nColumnas + 4).Value = 1
                       Cells(1, nColumnas + 5).Value = Cells(1, nColumnas + 5).Value + Cells(i, nColumnas + 4).Value
                    End If
               Else
                    
                     pos = i
                     
                     Cells(i - 1, nColumnas + 1).Value = "ultima actuación"
                        If Cells(1, nColumnas + 5).Value = 0 Then
                          Cells(i - 1, nColumnas + 5).Value = "ES DESCUENTO TODO"
                         Else
                          Cells(i - 1, nColumnas + 5).Value = "NO ES DESC"
                        End If
                        
                        If Cells(i - 1, nColumnas + 5).Value = "ES DESCUENTO TODO" Then
                          If pos > i - 1 Then
                           pos = pos2
                            For j = pos To i - 1
                            Cells(j, nColumnas + 2).Value = "descuento"
                            Next j
                          End If
                        End If
                        
                    Cells(1, nColumnas + 5).Value = 0
                    act = Cells(i, 14).Value
                    
                    If Cells(i, 9).Value = 2 Then
                     Cells(i, nColumnas + 4).Value = 0
                     Cells(1, nColumnas + 5).Value = Cells(1, nColumnas + 5).Value + Cells(i, nColumnas + 4).Value
                    Else
                        band = True
                       Cells(i, nColumnas + 2).Value = "ajuste en mas"
                       Cells(i, nColumnas + 4).Value = 1
                       Cells(1, nColumnas + 5).Value = Cells(1, nColumnas + 5).Value + Cells(i, nColumnas + 4).Value
                    End If
              End If
            Else
              Cells(i - 1, nColumnas + 1).Value = "ultima actuación"
              If Cells(1, nColumnas + 5).Value = 0 Then
               Cells(i - 1, nColumnas + 5).Value = "ES DESCUENTO TODO"
               Else
               Cells(i - 1, nColumnas + 5).Value = "NO ES DESC"
              End If
              
              If Cells(i - 1, nColumnas + 5).Value = "ES DESCUENTO TODO" Then
                 For j = pos To i - 1
                 Cells(j, nColumnas + 2).Value = "descuento"
                 Next j
              End If
                        
              Cells(1, nColumnas + 5).Value = 0
              act = Cells(i, 14).Value
                
                ultDoc = Cells(i, 5).Value
                importe = 0
                ultJur = Cells(i, 2).Value
                nombre = Cells(i, 7).Value
                pos = i
                pos2 = i
                i = i - 1
                
                
                Cells(i, nColumnas + 3).Value = "ultimo dni"
             
            End If
        End If
    Next i
  Cells(i - 1, nColumnas + 1).Value = "ultima actuación"
  Cells(i - 1, nColumnas + 3).Value = "ultimo dni"
   If Cells(1, nColumnas + 5).Value = 0 Then
     Cells(i - 1, nColumnas + 5).Value = "ES DESCUENTO TODO"
    Else
     Cells(i - 1, nColumnas + 5).Value = "NO ES DESC"
   End If
    
   If Cells(i - 1, nColumnas + 5).Value = "ES DESCUENTO TODO" Then
    For j = pos To i - 1
    Cells(j, nColumnas + 2).Value = "descuento"
    Next j
   End If
   
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Application.StatusBar = False
End Sub



