Attribute VB_Name = "Módulo1"
Sub INFORME_ACTUACIONES()
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
    
    For i = 2 To limite
      Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
      
      'Regresa el control a la hoja origen
       Sheets("ACT").Select
            
       pos2 = i
       act = Cells(i, 1).Value
       band = False
       cont = 0
       suma = 0
       
       For j = 2 To nFilasCont
       
       
        'Regresa el control a la hoja nueva
          Sheets("TOTALES X PERSONA").Select
          
         If Sheets("TOTALES X PERSONA").Cells(j, 7).Value = act Then
            'Cells(j, nColumnasCont).Value = Cells(j, nColumnasCont).Value + 1
            cont = cont + 1
            suma = suma + Cells(j, 6).Value
       
            'Regresa el control a la hoja origen
             Sheets("ACT").Select
                                                     
             'Cells(i, nColumnas).Value = cont
                                             
         End If
            Sheets("ACT").Select
            
             Cells(i, nColumnas).Value = cont
             If suma = 0 Then
              Cells(i, nColumnas + 1).Value = "CUMPLIDA"
              Else
            Cells(i, nColumnas + 1).Value = "NO CUMPLIDA"
                
             End If
              
       Next j
    
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
        
        For j = (i + 1) To nFilas
            If valorDoc = Cells(j, 2).Value Then
                Cells(j, 2).Interior.Color = RGB(153, 196, 195)
                Cells(i, 2).Interior.Color = RGB(153, 196, 195)
                Cells(j, nColumnas + 1).Value = "Repetido"
                
                If Cells(j, nColumnas + 2).Value = "" Then
                  Cells(j, nColumnas + 2).Value = i
                  Cells(i, nColumnas + 2).Value = i
                Else
                  Cells(i, nColumnas + 2).Value = Cells(j, nColumnas + 2).Value
                End If
                
                Cells(i, nColumnas + 1).Value = "Repetido"
            End If
        Next j
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
           
             
              For j = 2 To nFilasCont
              
              'Va la hoja del libro que voy a abrir
               Set wsContenido = wbContenido.Worksheets("INFORME ACT HISTORICO")

    
                 If wsContenido.Cells(j, 4).Value = juras Then
                    If wsContenido.Cells(j, 5).Value = año Then
                         If wsContenido.Cells(j, 6).Value = num Then
                            
                              
                            wsContenido.Cells(j, 13).Value = "EN INFORME NEG"
                                                 
                                               
                                             
                                                  
                            'libro que ya tengo abierto
                                                  
                            'Worksheets("ACT").Cells(i, 17).Value = wsContenido.Cells(j, 3).Value
                            Worksheets("ACT").Cells(i, nColumnas + 1).Value = "ENCONTRADA"
                           
                         End If
                    End If
                 End If
                
                
            Next j
          End If
    Next i
     MsgBox "Proceso exitoso"
     Application.StatusBar = False
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


