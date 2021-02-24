Attribute VB_Name = "Módulo1"
Sub Comparar_dif_Archivos()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i As Long
    Dim filaCopia As Long
    Dim nFilasError, inicio, limite As Integer
    Dim columnaCopia As Long
    Dim wbContenido As Workbook, _
        wsContenido As Excel.Worksheet


    
    'abro el libro con el que voy a comparar con el que tengo abierto
    Set wbContenido = Application.Workbooks.Open("D:\TRABAJO\CARGAR 2020\JUNIO 2020\APAREOS\JUR 06 GUARDIAS ABRIL 20.xlsx")
  
    
    'Activo el libro que estoy por abrir
    ThisWorkbook.Activate
    
    
    'va el nombre de la hoja del libro que voy a abrir
    Set wsContenido = wbContenido.Worksheets("Hoja1")
  
    
    'va el nombre de la hoja del libro que ya tengo abierto
    Sheets("CPTOS_J6_2020_5_1_1_Mes_actual").Select
   
    
    'Calculo el número de filas de la hoja actual del libro que ya tengo abierto
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Sheets("COM2005-08-PARA CARGAR-MT-17-06").Cells(1, 24).Value = "MENSUAL-HIST2 MAYO"
   ' Sheets("COM2005-08-PARA CARGAR-MT-17-06").Cells(1, 19).Value = "POSICIÓN-EN-999999"
    
    
    'Calculo el número de filas de la hoja de los cobrados
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    
    'el libro que voy a abrir
    Workbooks("JUR 06 GUARDIAS ABRIL 20.xlsx").Activate
    ' Sheets("COPIA 999999 DE SELECCION").Cells(1, 38).Value = "POS EN COM2005-08"
    
    
    'en la primer columna vacia la nombro para ver cuales son los que que se encuentran en cobrados
    'Cells(1, nColumnas + 1).Value = "IGUALES"
    
    
    limite = 10979
    For i = 2 To limite
    Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
       'libro que ya tengo abierto
        Workbooks("CPTOS_J6_2020_5_1_1_Mes_actual.csv").Activate
    
                pos2 = i
                dni = Cells(i, 2).Value
                'nombre = Cells(i, 3).Value
                'horas = Cells(i, 9).Value
                'cuofguardias = Cells(i, 1).Value
                'reaj = Cells(i, 9).Value
                'unidad = Cells(i, 10).Value
                'importe = Cells(i, 11).Value
                'vto = Cells(i, 12).Value
                'cuota = Cells(i, 13).Value
                'act = Cells(i, 14).Value
                'fila = Cells(i, 32).Value
                'cuotatotal = Cells(i, 18).Value
                'totalcuota = Cells(i, 28).Value
                'habilito = Cells(i, 29).Value
                'partir = Cells(i, 30).Value
                'coupend = Cells(i, 31).Value
                'esta = Cells(i, 33).Value
              For j = 2 To 1459
                 'el libro que voy a abrir
                Workbooks("JUR 06 GUARDIAS ABRIL 20.xlsx").Activate
                
                 If Sheets("Hoja1").Cells(j, 5).Value = dni Then
                    'If Sheets("Hoja1").Cells(j, 6).Value = nombre Then
                         'If Sheets("Hoja1").Cells(j, 3).Value = horas Then
                            'If Sheets("Hoja1").Cells(j, 8).Value = cuofguardias Then
                           
                               
                                    tipoprof = Sheets("Hoja1").Cells(j, 7).Value
                                    horas = Sheets("Hoja1").Cells(j, 9).Value
                                    cuofguardias = Sheets("Hoja1").Cells(j, 1).Value
                                    
                                    pos = j
                                    Sheets("Hoja1").Cells(j, 15).Value = pos2
                                                
                                                    
                                                    'libro que ya tengo abierto
                                                    Workbooks("CPTOS_J6_2020_5_1_1_Mes_actual.csv").Activate
                                                    Sheets("CPTOS_J6_2020_5_1_1_Mes_actual").Cells(i, 25).Value = tipoprof
                                                    Sheets("CPTOS_J6_2020_5_1_1_Mes_actual").Cells(i, 26).Value = horas
                                                    Sheets("CPTOS_J6_2020_5_1_1_Mes_actual").Cells(i, 27).Value = cuofguardias
                                                    
                                                     If Sheets("CPTOS_J6_2020_5_1_1_Mes_actual").Cells(i, 29).Value = "" Then
                                                         Sheets("CPTOS_J6_2020_5_1_1_Mes_actual").Cells(i, 29).Value = pos
                                                        
                                                     End If
                                             
                              
                             'End If
                         'End If
                    'End If
                 End If
               
                
            Next j
   
    Next i
     MsgBox "Proceso exitoso"
     Application.StatusBar = False
     
End Sub












