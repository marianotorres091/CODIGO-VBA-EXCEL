Attribute VB_Name = "Módulo1"
Sub Totales_Persona_Deuda_Total_PASO1()
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
    
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Total x Persona_DT"
    Application.DisplayAlerts = True
    
    Set wsTotal = Worksheets("Total x Persona_DT")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    MsgBox "Debe estar ordenado por DNI.", , "Atención!!"
    
    'Encabezado Hoja Totales
    wsTotal.Cells(1, 1).Value = "JUR"
    wsTotal.Cells(1, 2).Value = "DNI"
    wsTotal.Cells(1, 3).Value = "NOMBRE"
    wsTotal.Cells(1, 4).Value = "DEUDA TOTAL"
    wsTotal.Cells(1, 5).Value = "ACTUACION"
    wsTotal.Cells(1, 6).Value = "OPERADOR"
    
    filaTotal = 2
    colTotal = 5
    
    ultDoc = Cells(2, 5).Value
    importe = 0
    ultJur = Cells(2, 2).Value
    nombre = Cells(2, 7).Value
    act = Cells(2, 15).Value
    operador = Cells(2, 18).Value
    
    limite = nFilas
    
    For i = 2 To limite
    Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
        If Cells(i, 8).Value < 400 Then
            If ultDoc = Cells(i, 5).Value And act = Cells(i, 15).Value Then
              operador = Cells(i, 18).Value
                If Cells(i, 9).Value = 2 Then
                    importe = importe - Cells(i, 11).Value
                Else
                    importe = importe + Cells(i, 11).Value
                End If
            Else
                wsTotal.Cells(filaTotal, 1).Value = ultJur
                wsTotal.Cells(filaTotal, 2).Value = ultDoc
                wsTotal.Cells(filaTotal, 3).Value = nombre
                wsTotal.Cells(filaTotal, 4).Value = importe
                
                If wsTotal.Cells(filaTotal, 5).Value = "" Then
                 wsTotal.Cells(filaTotal, 5).Value = act
                 Else
                 wsTotal.Cells(filaTotal, colTotal + 1).Value = act
                End If
                
                wsTotal.Cells(filaTotal, 6).Value = operador
                
                filaTotal = filaTotal + 1
                
                ultDoc = Cells(i, 5).Value
                importe = 0
                ultJur = Cells(i, 2).Value
                nombre = Cells(i, 7).Value
                act = Cells(i, 15).Value
                i = i - 1
            End If
        Else
           If ultDoc = Cells(i, 5).Value And act = Cells(i, 15).Value Then
              operador = Cells(i, 18).Value
                If Cells(i, 9).Value = 1 Then
                    importe = importe - Cells(i, 11).Value
                Else
                    importe = importe + Cells(i, 11).Value
                End If
            Else
                wsTotal.Cells(filaTotal, 1).Value = ultJur
                wsTotal.Cells(filaTotal, 2).Value = ultDoc
                wsTotal.Cells(filaTotal, 3).Value = nombre
                wsTotal.Cells(filaTotal, 4).Value = importe
                
                If wsTotal.Cells(filaTotal, 5).Value = "" Then
                 wsTotal.Cells(filaTotal, 5).Value = act
                 Else
                 wsTotal.Cells(filaTotal, colTotal + 1).Value = act
                End If
                
                wsTotal.Cells(filaTotal, 6).Value = operador
                
                filaTotal = filaTotal + 1
                
                ultDoc = Cells(i, 5).Value
                importe = 0
                ultJur = Cells(i, 2).Value
                nombre = Cells(i, 7).Value
                act = Cells(i, 15).Value
                i = i - 1
            End If
        End If
    Next i
    wsTotal.Cells(filaTotal, 1).Value = ultJur
    wsTotal.Cells(filaTotal, 2).Value = ultDoc
    wsTotal.Cells(filaTotal, 3).Value = nombre
    wsTotal.Cells(filaTotal, 4).Value = importe
    
    If wsTotal.Cells(filaTotal, 5).Value = "" Then
        wsTotal.Cells(filaTotal, 5).Value = act
      Else
        wsTotal.Cells(filaTotal, colTotal + 1).Value = act
    End If
    
    wsTotal.Cells(filaTotal, 6).Value = operador
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Application.StatusBar = False
End Sub

Sub Total_Liquidado_Persona_ult_liquidacion_PASO2()
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
    
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Total x Persona_LIQ"
    Application.DisplayAlerts = True
    
    Set wsTotal = Worksheets("Total x Persona_LIQ")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    MsgBox "Debe estar ordenado por DNI.", , "Atención!!"
    
    'Encabezado Hoja Totales
    wsTotal.Cells(1, 1).Value = "JUR"
    wsTotal.Cells(1, 2).Value = "DNI"
    wsTotal.Cells(1, 3).Value = "NOMBRE"
    wsTotal.Cells(1, 4).Value = "LIQUIDADO"
    wsTotal.Cells(1, 5).Value = "ACTUACION"
    wsTotal.Cells(1, 6).Value = "ULT LIQUIDACION"
    filaTotal = 2
    colTotal = 5
    
    ultDoc = Cells(2, 5).Value
    importe = 0
    ultJur = Cells(2, 2).Value
    nombre = Cells(2, 7).Value
    act = Cells(2, 15).Value
    ultliq = 0

    limite = nFilas
    
    For i = 2 To limite
    Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
    
        If Cells(i, 8).Value < 400 Then
            If ultDoc = Cells(i, 5).Value And act = Cells(i, 15).Value Then
             If Cells(i, 14).Value <> "" Then
               liq = Cells(i, 14).Value
              Select Case liq
                     Case Is = "MEN012020"
                      liq = 1
                     Case Is = "COM0120-08"
                      liq = 2
                     Case Is = "MEN022020"
                      liq = 3
                     Case Is = "COM0220-08"
                      liq = 4
                     Case Is = "MEN032020"
                      liq = 5
                     Case Is = "COM0320-08"
                      liq = 6
                     Case Is = "MEN042020"
                      liq = 7
                     Case Is = "COM0420-08"
                      liq = 8
                     Case Is = "MEN052020"
                      liq = 9
                     Case Is = "COM0520-09"
                      liq = 10
                     Case Is = "MEN062020"
                      liq = 11
                     Case Is = "COM0620-08"
                      liq = 12
                     Case Is = "MEN072020"
                      liq = 13
                     Case Is = "COM0720-09"
                      liq = 14
                     Case Is = "MEN082020"
                      liq = 15
                     Case Is = "COM0820-12"
                      liq = 16
                     Case Is = "MEN092020"
                      liq = 17
                     Case Is = "COM0920-16"
                      liq = 18
                     Case Is = "MEN102020"
                      liq = 19
                     Case Is = "COM1020-11"
                      liq = 20
                     Case Is = "MEN112020"
                      liq = 21
                     Case Else
                      liq = 0
                    End Select
                    
                   If ultliq < liq Then
                     ultliq = liq
                   End If
                                     
                  
                       
                     
                If Cells(i, 9).Value = 2 Then
                    importe = importe - Cells(i, 11).Value
                Else
                    importe = importe + Cells(i, 11).Value
                End If
             End If
            Else
                wsTotal.Cells(filaTotal, 1).Value = ultJur
                wsTotal.Cells(filaTotal, 2).Value = ultDoc
                wsTotal.Cells(filaTotal, 3).Value = nombre
                wsTotal.Cells(filaTotal, 4).Value = importe
                If wsTotal.Cells(filaTotal, 5).Value = "" Then
                 wsTotal.Cells(filaTotal, 5).Value = act
                 Else
                 wsTotal.Cells(filaTotal, colTotal + 1).Value = act
                End If
                
                If importe = 0 Then
                wsTotal.Cells(filaTotal, 6).Value = "no posee liquidacion"
                Else
                Select Case ultliq
                     Case Is = 1
                      ultliq = "MEN012020"
                     Case Is = 2
                      ultliq = "COM0120 - 8"
                     Case Is = 3
                      ultliq = "MEN022020"
                     Case Is = 4
                      ultliq = "COM0220-08"
                     Case Is = 5
                      ultliq = "MEN032020"
                     Case Is = 6
                      ultliq = "COM0320-08"
                     Case Is = 7
                      ultliq = "MEN042020"
                     Case Is = 8
                      ultliq = "COM0420-08"
                     Case Is = 9
                      ultliq = "MEN052020"
                     Case Is = 10
                      ultliq = "COM0520-09"
                     Case Is = 11
                      ultliq = "MEN062020"
                     Case Is = 12
                      ultliq = "COM0620-08"
                     Case Is = 13
                      ultliq = "MEN072020"
                     Case Is = 14
                      ultliq = "COM0720-09"
                     Case Is = 15
                      ultliq = "MEN082020"
                     Case Is = 16
                      ultliq = "COM0820-12"
                     Case Is = 17
                      ultliq = "MEN092020"
                     Case Is = 18
                      ultliq = "COM0920-16"
                     Case Is = 19
                      ultliq = "MEN102020"
                     Case Is = 20
                      ultliq = "COM1020-11"
                     Case Is = 21
                      ultliq = "MEN112020"
                     Case Is = 0
                      ultliq = "ver"
                    End Select
                wsTotal.Cells(filaTotal, 6).Value = ultliq
                End If
                
                filaTotal = filaTotal + 1
                ultliq = 0
                ultDoc = Cells(i, 5).Value
                importe = 0
                ultJur = Cells(i, 2).Value
                nombre = Cells(i, 7).Value
                act = Cells(i, 15).Value
                liq = Cells(i, 14).Value
                i = i - 1
            End If
        Else
           If ultDoc = Cells(i, 5).Value And act = Cells(i, 15).Value Then
             If Cells(i, 14).Value <> "" Then
             
             liq = Cells(i, 14).Value
             
             Select Case liq
                     Case Is = "MEN012020"
                      liq = 1
                     Case Is = "COM0120-08"
                      liq = 2
                     Case Is = "MEN022020"
                      liq = 3
                     Case Is = "COM0220-08"
                      liq = 4
                     Case Is = "MEN032020"
                      liq = 5
                     Case Is = "COM0320-08"
                      liq = 6
                     Case Is = "MEN042020"
                      liq = 7
                     Case Is = "COM0420-08"
                      liq = 8
                     Case Is = "MEN052020"
                      liq = 9
                     Case Is = "COM0520-09"
                      liq = 10
                     Case Is = "MEN062020"
                      liq = 11
                     Case Is = "COM0620-08"
                      liq = 12
                     Case Is = "MEN072020"
                      liq = 13
                     Case Is = "COM0720-09"
                      liq = 14
                     Case Is = "MEN082020"
                      liq = 15
                     Case Is = "COM0820-12"
                      liq = 16
                     Case Is = "MEN092020"
                      liq = 17
                     Case Is = "COM0920-16"
                      liq = 18
                     Case Is = "MEN102020"
                      liq = 19
                     Case Is = "COM1020-11"
                      liq = 20
                     Case Is = "MEN112020"
                      liq = 21
                     Case Else
                      liq = 0
                    End Select
                    
                   If ultliq < liq Then
                     ultliq = liq
                   End If
                   
                If Cells(i, 9).Value = 1 Then
                    importe = importe - Cells(i, 11).Value
                Else
                    importe = importe + Cells(i, 11).Value
                End If
              End If
            Else
                wsTotal.Cells(filaTotal, 1).Value = ultJur
                wsTotal.Cells(filaTotal, 2).Value = ultDoc
                wsTotal.Cells(filaTotal, 3).Value = nombre
                wsTotal.Cells(filaTotal, 4).Value = importe
                If wsTotal.Cells(filaTotal, 5).Value = "" Then
                 wsTotal.Cells(filaTotal, 5).Value = act
                 Else
                 wsTotal.Cells(filaTotal, colTotal + 1).Value = act
                End If
                
                If importe = 0 Then
                wsTotal.Cells(filaTotal, 6).Value = "no posee liquidacion"
                Else
                
                Select Case ultliq
                     Case Is = 1
                      ultliq = "MEN012020"
                     Case Is = 2
                      ultliq = "COM0120 - 8"
                     Case Is = 3
                      ultliq = "MEN022020"
                     Case Is = 4
                      ultliq = "COM0220-08"
                     Case Is = 5
                      ultliq = "MEN032020"
                     Case Is = 6
                      ultliq = "COM0320-08"
                     Case Is = 7
                      ultliq = "MEN042020"
                     Case Is = 8
                      ultliq = "COM0420-08"
                     Case Is = 9
                      ultliq = "MEN052020"
                     Case Is = 10
                      ultliq = "COM0520-09"
                     Case Is = 11
                      ultliq = "MEN062020"
                     Case Is = 12
                      ultliq = "COM0620-08"
                     Case Is = 13
                      ultliq = "MEN072020"
                     Case Is = 14
                      ultliq = "COM0720-09"
                     Case Is = 15
                      ultliq = "MEN082020"
                     Case Is = 16
                      ultliq = "COM0820-12"
                     Case Is = 17
                      ultliq = "MEN092020"
                     Case Is = 18
                      ultliq = "COM0920-16"
                     Case Is = 19
                      ultliq = "MEN102020"
                     Case Is = 20
                      ultliq = "COM1020-11"
                     Case Is = 21
                      ultliq = "MEN112020"
                     Case Is = 0
                      ultliq = "ver"
                    End Select
                    
                  wsTotal.Cells(filaTotal, 6).Value = ultliq
                End If
                
                filaTotal = filaTotal + 1
                
                ultliq = 0
                ultDoc = Cells(i, 5).Value
                importe = 0
                ultJur = Cells(i, 2).Value
                nombre = Cells(i, 7).Value
                act = Cells(i, 15).Value
                liq = Cells(i, 14).Value
                i = i - 1
            End If
        End If
   
    Next i
    wsTotal.Cells(filaTotal, 1).Value = ultJur
    wsTotal.Cells(filaTotal, 2).Value = ultDoc
    wsTotal.Cells(filaTotal, 3).Value = nombre
    wsTotal.Cells(filaTotal, 4).Value = importe
   If wsTotal.Cells(filaTotal, 5).Value = "" Then
    wsTotal.Cells(filaTotal, 5).Value = act
    Else
    wsTotal.Cells(filaTotal, colTotal + 1).Value = act
   End If
   
    If importe = 0 Then
        wsTotal.Cells(filaTotal, 6).Value = "no posee liquidacion"
     Else
     
     Select Case ultliq
                     Case Is = 1
                      ultliq = "MEN012020"
                     Case Is = 2
                      ultliq = "COM0120 - 8"
                     Case Is = 3
                      ultliq = "MEN022020"
                     Case Is = 4
                      ultliq = "COM0220-08"
                     Case Is = 5
                      ultliq = "MEN032020"
                     Case Is = 6
                      ultliq = "COM0320-08"
                     Case Is = 7
                      ultliq = "MEN042020"
                     Case Is = 8
                      ultliq = "COM0420-08"
                     Case Is = 9
                      ultliq = "MEN052020"
                     Case Is = 10
                      ultliq = "COM0520-09"
                     Case Is = 11
                      ultliq = "MEN062020"
                     Case Is = 12
                      ultliq = "COM0620-08"
                     Case Is = 13
                      ultliq = "MEN072020"
                     Case Is = 14
                      ultliq = "COM0720-09"
                     Case Is = 15
                      ultliq = "MEN082020"
                     Case Is = 16
                      ultliq = "COM0820-12"
                     Case Is = 17
                      ultliq = "MEN092020"
                     Case Is = 18
                      ultliq = "COM0920-16"
                     Case Is = 19
                      ultliq = "MEN102020"
                     Case Is = 20
                      ultliq = "COM1020-11"
                     Case Is = 21
                      ultliq = "MEN112020"
                     Case Is = 0
                      ultliq = "ver"
                    End Select
                    
        wsTotal.Cells(filaTotal, 6).Value = ultliq
    End If
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Application.StatusBar = False
End Sub

Sub Totales_Persona_Deuda_Pendiente_PASO3()
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
    
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Total x Persona_DP"
    Application.DisplayAlerts = True
    
    Set wsTotal = Worksheets("Total x Persona_DP")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    MsgBox "Debe estar ordenado por DNI.", , "Atención!!"
    
    'Encabezado Hoja Totales
    wsTotal.Cells(1, 1).Value = "JUR"
    wsTotal.Cells(1, 2).Value = "DNI"
    wsTotal.Cells(1, 3).Value = "NOMBRE"
    wsTotal.Cells(1, 4).Value = "DEUDA PENDIENTE"
    wsTotal.Cells(1, 5).Value = "ACTUACION"
   
    filaTotal = 2
    colTotal = 5
    
    ultDoc = Cells(2, 5).Value
    importe = 0
    ultJur = Cells(2, 2).Value
    nombre = Cells(2, 7).Value
    act = Cells(2, 15).Value
   

    limite = nFilas
    
    For i = 2 To limite
    Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
    
        If Cells(i, 8).Value < 400 Then
            If ultDoc = Cells(i, 5).Value And act = Cells(i, 15).Value Then
             If Cells(i, 14).Value = "" Then
              
                If Cells(i, 9).Value = 2 Then
                    importe = importe - Cells(i, 11).Value
                Else
                    importe = importe + Cells(i, 11).Value
                End If
             End If
            Else
                wsTotal.Cells(filaTotal, 1).Value = ultJur
                wsTotal.Cells(filaTotal, 2).Value = ultDoc
                wsTotal.Cells(filaTotal, 3).Value = nombre
                wsTotal.Cells(filaTotal, 4).Value = importe
                If wsTotal.Cells(filaTotal, 5).Value = "" Then
                 wsTotal.Cells(filaTotal, 5).Value = act
                 Else
                 wsTotal.Cells(filaTotal, colTotal + 1).Value = act
                End If
                
                filaTotal = filaTotal + 1
                
                ultDoc = Cells(i, 5).Value
                importe = 0
                ultJur = Cells(i, 2).Value
                nombre = Cells(i, 7).Value
                act = Cells(i, 15).Value
               
                i = i - 1
            End If
        Else
           If ultDoc = Cells(i, 5).Value And act = Cells(i, 15).Value Then
             If Cells(i, 14).Value = "" Then
         
                If Cells(i, 9).Value = 1 Then
                    importe = importe - Cells(i, 11).Value
                Else
                    importe = importe + Cells(i, 11).Value
                End If
              End If
            Else
                wsTotal.Cells(filaTotal, 1).Value = ultJur
                wsTotal.Cells(filaTotal, 2).Value = ultDoc
                wsTotal.Cells(filaTotal, 3).Value = nombre
                wsTotal.Cells(filaTotal, 4).Value = importe
                If wsTotal.Cells(filaTotal, 5).Value = "" Then
                 wsTotal.Cells(filaTotal, 5).Value = act
                 Else
                 wsTotal.Cells(filaTotal, colTotal + 1).Value = act
                End If
                
         
                filaTotal = filaTotal + 1
                
                ultDoc = Cells(i, 5).Value
                importe = 0
                ultJur = Cells(i, 2).Value
                nombre = Cells(i, 7).Value
                act = Cells(i, 15).Value
               
                i = i - 1
            End If
        End If
   
    Next i
    wsTotal.Cells(filaTotal, 1).Value = ultJur
    wsTotal.Cells(filaTotal, 2).Value = ultDoc
    wsTotal.Cells(filaTotal, 3).Value = nombre
    wsTotal.Cells(filaTotal, 4).Value = importe
   If wsTotal.Cells(filaTotal, 5).Value = "" Then
    wsTotal.Cells(filaTotal, 5).Value = act
    Else
    wsTotal.Cells(filaTotal, colTotal + 1).Value = act
   End If
   
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Application.StatusBar = False
End Sub



