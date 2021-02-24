Attribute VB_Name = "Módulo1"
Sub DETERMINAR_ACT_CUMPLIDAS()
    Dim rango As Range
    Dim rangoCont As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim nColumnasCont As Long
    Dim i As Long
    Dim filaCopia As Long
  
    'Regresa el control a la hoja de origen
     Sheets("ACT").Select
   
    
    'Calcular el número de filas de la hoja actual
     Set rango = ActiveSheet.UsedRange
     nFilas = rango.Rows.Count
     nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
     Sheets("TOTALES X PERSONA").Select
     Set rangoCont = ActiveSheet.UsedRange
     nFilasCont = rangoCont.Rows.Count
     nColumnasCont = rangoCont.Columns.Count
    
    limite = nFilas
    nColumnas = nColumnas + 1
    nColumnasCont = nColumnasCont + 1
    
    Sheets("ACT").Cells(1, nColumnas).Value = "CANT AGENTES"
    Sheets("ACT").Cells(1, nColumnas + 1).Value = "ESTADO"
    
    For i = 2 To limite
      Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
      
      'Regresa el control a la hoja origen
       Sheets("ACT").Select
            
       pos2 = i
       act = Cells(i, 1).Value
       band = False
       cont = 0
       suma = 0
       
       For J = 2 To nFilasCont
       
       
        'Regresa el control a la hoja nueva
          Sheets("TOTALES X PERSONA").Select
          
         If Sheets("TOTALES X PERSONA").Cells(J, 7).Value = act Then
            'Cells(j, nColumnasCont).Value = Cells(j, nColumnasCont).Value + 1
            cont = cont + 1
            suma = suma + Cells(J, 6).Value
       
            'Regresa el control a la hoja origen
             Sheets("ACT").Select
                                                     
             Cells(i, nColumnas).Value = cont
                                             
         End If
            Sheets("ACT").Select
            
             Cells(i, nColumnas).Value = cont
             If suma = 0 Then
              Cells(i, nColumnas + 1).Value = "CUMPLIDA"
              Else
            Cells(i, nColumnas + 1).Value = "NO CUMPLIDA"
                
             End If
              
       Next J
    
    Next i
    
    'Regresa el control a la hoja de origen
     Sheets("ACT").Select
    
     MsgBox "Proceso exitoso"
     Application.StatusBar = False

End Sub

Sub Registros_duplicados()
    Dim nFilas As Long
    Dim nColumnas As Long
    'Dim i As Integer
    Dim rango As Range
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
   
    limite = (nFilas - 1)
    For i = 2 To limite
     Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
     If Cells(i, nColumnas + 1).Value <> "Repetido" Then
     
        'valorjur = Cells(i, 2).Value
        'valorescalafon = Cells(i, 3).Value
        valorDoc = Cells(i, 2).Value
        'valorCouc = Cells(i, 8).Value
        'valorRj = Cells(i, 10).Value
        'valorunidad = Cells(i, 11).Value
        'valorimporte = Cells(i, 12).Value
        'valorVto = Cells(i, 15).Value
        'cuota = Cells(i, 13).Value
        
        For J = (i + 1) To nFilas
            If valorDoc = Cells(J, 2).Value Then
                Cells(J, 2).Interior.Color = RGB(153, 196, 195)
                Cells(i, 2).Interior.Color = RGB(153, 196, 195)
                Cells(J, nColumnas + 1).Value = "Repetido"
                
                If Cells(J, nColumnas + 2).Value = "" Then
                  Cells(J, nColumnas + 2).Value = i
                  Cells(i, nColumnas + 2).Value = i
                Else
                  Cells(i, nColumnas + 2).Value = Cells(J, nColumnas + 2).Value
                End If
                
                Cells(i, nColumnas + 1).Value = "Repetido"
            End If
        Next J
      End If
    Next i
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Application.StatusBar = False
End Sub
Sub SepararApeyNom()
Dim celda As Range       'celda que contiene el texto
Dim i As Integer
Dim n As Integer         'número de palabras encontradas
Dim palabras() As String 'arreglo que almacenará las palabras separadas
Dim separador As String  'separador de cada palabra
Dim texto As String      'almacena el texto a separar
Dim rango As Range
Dim cell As Range, AreaTotrim As Range

'Calculo el número de filas de la hoja actual del libro que ya tengo abierto
     Set rango = ActiveSheet.UsedRange
     nFilas = rango.Rows.Count
     
     rangoTemp = "C1:C" & nFilas
     
     Set AreaTotrim = Worksheets("ACT").Range(rangoTemp)
     
    
    'definir el separador de palabras
    separador = "-" 'espacio en blanco
    
    'Ciclo para recorrer los renglones
    For Each celda In Selection
        texto = celda.Value
        
        'Separación del texto en palabras:
        palabras = Split(texto, separador)
        
        'La función UBound devuelve índice final/mayor del arreglo
        'El índice en el arreglo se inicia con cero
        n = UBound(palabras)
        
        'Ciclo para colocar cada palabra en una columna diferente
        For i = 0 To n
            celda.Offset(0, i + 1) = palabras(i)
        Next i
 
    Next celda
     
    'Elimino el primer caracter en blanco de los nombres
    For Each cell In AreaTotrim
      cell = Trim(cell)
    Next cell
  
    MsgBox "Proceso exitoso"
    
End Sub

Sub Comprar_Archivos_Diferentes()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim nColumnasCont As Long
    Dim i As Long
    Dim filaCopia As Long
    Dim nFilasError As Integer
    Dim columnaCopia As Long
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet


    'Indicar el libro de excel CONTENIDO y control de errores
    contenido = InputBox("Ingrese el nombre del archivo:", "Abrir", "Archivo.xlsx")
    If contenido <> "" Then
       ' On Error GoTo ControlErrorOpen
        Set wbContenido = Workbooks.Open(ActiveWorkbook.Path & "\" & contenido)
    Else
        Exit Sub
    End If
    
    'Activar este libro
    ThisWorkbook.Activate
    
    Application.DisplayAlerts = False
    'Worksheets.Add
    'ActiveSheet.Name = "Errores"
    Application.DisplayAlerts = True
    'Set wsError = Worksheets("Errores")
    'Va la Hoja del Libro que se va a Abrir
    Set wsContenido = wbContenido.Worksheets("INFORME ACT HISTORICO")
    
    'Regresa el control a la hoja de origen
    Sheets("ACT").Select
    
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    nColumnasCont = rangoCont.Columns.Count
    
    limite = nFilas
    For i = 2 To limite
    Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
       'libro que ya tengo abierto
                pos2 = i
            If Cells(i, 2).Value <> "" And Cells(i, 3).Value <> "" And Cells(i, 4).Value <> "" Then
            
                juras = Cells(i, 2).Value
                año = Cells(i, 3).Value
                num = Cells(i, 4).Value
           
             
              For J = 2 To nFilasCont
              
              'Va la hoja del libro que voy a abrir
               Set wsContenido = wbContenido.Worksheets("INFORME ACT HISTORICO")

    
                 If wsContenido.Cells(J, 4).Value = juras Then
                    If wsContenido.Cells(J, 5).Value = año Then
                         If wsContenido.Cells(J, 6).Value = num Then
                            
                              
                            wsContenido.Cells(J, 13).Value = "EN INFORME NEG"
                                                 
                                               
                                             
                                                  
                            'libro que ya tengo abierto
                                                  
                            'Worksheets("ACT").Cells(i, 17).Value = wsContenido.Cells(j, 3).Value
                            Worksheets("ACT").Cells(i, nColumnas + 1).Value = "ENCONTRADA"
                           
                         End If
                    End If
                 End If
                
                
            Next J
          End If
    Next i
     MsgBox "Proceso exitoso"
     Application.StatusBar = False
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub

Sub COPIAR_ULT_LIQUIDACION()
    Dim rango As Range
    Dim rangoCont As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim nColumnasCont As Long
    Dim i As Long
    Dim filaCopia As Long
  
     MsgBox "COPIA LA ULT LIQUIDACION DE LA HOJA *TOTALES X PERSONA* A LA HOJA *ACTUACIONES*"
     MsgBox "LA HOJA *TOTALES X PERSONA* DEBE ESTAR ORDENADA POR *ACTUACIÓN*"
     
    'Regresa el control a la hoja de origen
     Sheets("ACT").Select
   
    
    'Calcular el número de filas de la hoja actual
     Set rango = ActiveSheet.UsedRange
     nFilas = rango.Rows.Count
     nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
     Sheets("TOTALES X PERSONA").Select
     Set rangoCont = ActiveSheet.UsedRange
     nFilasCont = rangoCont.Rows.Count
     nColumnasCont = rangoCont.Columns.Count
    
    limite = nFilasCont
  
    
    Sheets("ACT").Cells(1, nColumnas + 1).Value = "ULT LIQUIDACION"
    Sheets("ACT").Cells(1, nColumnas + 2).Value = "OPERADOR"
    
    
 For i = 2 To limite
    Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
    
    operador = Sheets("TOTALES X PERSONA").Cells(i, 9).Value
    
    liq = Sheets("TOTALES X PERSONA").Cells(i, 8).Value
          
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
                     Case Is = "COM1120-09"
                      liq = 22
                     Case Is = "MEN122020"
                      liq = 23
                     Case Is = "no posee liquidacion"
                      liq = 0
                     Case Else
                      liq = -1
                    End Select
                    
        mayor1 = liq

       act = Sheets("TOTALES X PERSONA").Cells(i, 7).Value
       

       For t = 2 To nFilasCont
        
       
        'Regresa el control a la hoja nueva
          Sheets("TOTALES X PERSONA").Select
          
          'CALCULO DEL PRIMER MAYOR
          
         If Sheets("TOTALES X PERSONA").Cells(t, 7).Value = act Then
          oper = Sheets("TOTALES X PERSONA").Cells(t, 9).Value
          liq = Sheets("TOTALES X PERSONA").Cells(t, 8).Value
          
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
                              Case Is = "COM1120-09"
                               liq = 22
                              Case Is = "MEN122020"
                               liq = 23
                              Case Is = "no posee liquidacion"
                               liq = 0
                              Case Else
                               liq = -1
                             End Select

                 If liq > mayor1 Then
                    mayor1 = liq
                    
                    For Z = 2 To nFilas
                     If act = Sheets("ACT").Cells(Z, 1).Value Then
                     
                       ultliq = mayor1
                     
                            Select Case ultliq
                             Case Is = 1
                              ultliq = "MEN012020"
                             Case Is = 2
                              ultliq = "COM0120-08"
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
                             Case Is = 22
                              ultliq = "COM1120-09"
                             Case Is = 23
                              ultliq = "MEN122020"
                             Case Is = 0
                              ultliq = "no posee liquidacion"
                             Case Is = -1
                              ultliq = "ver"
                            End Select
                            
                      Sheets("ACT").Cells(Z, nColumnas + 1).Value = ultliq
                      Sheets("ACT").Cells(Z, nColumnas + 2).Value = oper
                     End If
                    Next Z
                    posmayor1 = J
                    
                   Else
                  
                    If liq < mayor1 Then
                    
                   
                   
                     Else
                     
                       If liq = mayor1 Then
                         
                    
                            For h = 2 To nFilas
                             If act = Sheets("ACT").Cells(h, 1).Value Then
                             
                               ultliq = mayor1
                             
                                    Select Case ultliq
                                     Case Is = 1
                                      ultliq = "MEN012020"
                                     Case Is = 2
                                      ultliq = "COM0120-08"
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
                                     Case Is = 22
                                      ultliq = "COM1120-09"
                                     Case Is = 23
                                      ultliq = "MEN122020"
                                     Case Is = 0
                                      ultliq = "no posee liquidacion"
                                     Case Is = -1
                                      ultliq = "ver"
                                    End Select
                                    
                              Sheets("ACT").Cells(h, nColumnas + 1).Value = ultliq
                              Sheets("ACT").Cells(h, nColumnas + 2).Value = operador
                             End If
                            Next h
                       End If
                    End If
                   
                 End If
                            
         End If
            
              
       Next t
    
 Next i
    
    'Regresa el control a la hoja de origen
     Sheets("ACT").Select
    
     MsgBox "Proceso exitoso"
     Application.StatusBar = False

End Sub



