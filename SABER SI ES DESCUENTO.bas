Attribute VB_Name = "Módulo1"
Sub Saber_si_Persona_es_DESCUENTO()
    Dim total_mov As Long
    Dim total_dni As Long
    Dim total_actuaciones As Long
    Dim nFilas As Long
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
    
    MsgBox "Debe estar ordenado por DNI.", , "Atención!!"
    
    'INICIALIZAMOS VARIABLES
    act = Cells(2, 14).Value
    ultDoc = Cells(2, 5).Value
    importe = 0
    ultJur = Cells(2, 2).Value
    nombre = Cells(2, 7).Value
    limite = nFilas
    Cells(1, 29).Value = 0
    pos = 2
    
    For i = 2 To limite
    Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
        If Cells(i, 4).Value < 350 Then
            If ultDoc = Cells(i, 5).Value Then
               If act = Cells(i, 14).Value Then
               
                    If Cells(i, 9).Value = 2 Then
                     Cells(i, 28).Value = 0
                     Cells(1, 29).Value = Cells(1, 29).Value + Cells(i, 28).Value
                    Else
                        band = True
                       Cells(i, 26).Value = "ajuste en mas"
                       Cells(i, 28).Value = 1
                       Cells(1, 29).Value = Cells(1, 29).Value + Cells(i, 28).Value
                    End If
               Else
                     Cells(i - 1, 25).Value = "ultima actuación"
                        If Cells(1, 29).Value = 0 Then
                          Cells(i - 1, 29).Value = "ES DESCUENTO TODO"
                         Else
                          Cells(i - 1, 29).Value = "NO ES DESC"
                        End If
                        
                        If Cells(i - 1, 29).Value = "ES DESCUENTO TODO" Then
                         For j = pos To i - 1
                         Cells(j, 26).Value = "descuento"
                         Next j
                        End If
                        
                    Cells(1, 29).Value = 0
                    act = Cells(i, 14).Value
                    
                    If Cells(i, 9).Value = 2 Then
                     Cells(i, 28).Value = 0
                     Cells(1, 29).Value = Cells(1, 29).Value + Cells(i, 28).Value
                    Else
                        band = True
                       Cells(i, 26).Value = "ajuste en mas"
                       Cells(i, 28).Value = 1
                       Cells(1, 29).Value = Cells(1, 29).Value + Cells(i, 28).Value
                    End If
              End If
            Else
              
              If Cells(1, 29).Value = 0 Then
               Cells(i - 1, 29).Value = "ES DESCUENTO TODO"
               Else
               Cells(i - 1, 29).Value = "NO ES DESC"
              End If
              
              If Cells(i - 1, 29).Value = "ES DESCUENTO TODO" Then
                 For j = pos To i - 1
                 Cells(j, 26).Value = "descuento"
                 Next j
              End If
                        
              Cells(1, 29).Value = 0
              act = Cells(i, 14).Value
                
                ultDoc = Cells(i, 5).Value
                importe = 0
                ultJur = Cells(i, 2).Value
                nombre = Cells(i, 7).Value
                i = i - 1
                pos = i
                
                Cells(i, 27).Value = "ultimo dni"
             
            End If
        End If
    Next i
  Cells(i - 1, 27).Value = "ultimo dni"
   If Cells(1, 29).Value = 0 Then
     Cells(i - 1, 29).Value = "ES DESCUENTO TODO"
    Else
     Cells(i - 1, 29).Value = "NO ES DESC"
   End If
    
   If Cells(i - 1, 29).Value = "ES DESCUENTO TODO" Then
    For j = pos To i - 1
    Cells(j, 26).Value = "descuento"
    Next j
   End If
   
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Application.StatusBar = False
End Sub

