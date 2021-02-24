Attribute VB_Name = "Módulo1"
Sub Verificar_si_existe_dni_en_HistoricoGuardias()
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
    Application.DisplayAlerts = True
    Set wsContenido = wbContenido.Worksheets("HISTORICO")
    
    'Regresa el control a la hoja de origen
    Sheets("VERIF TIPOPROF").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    'Calcular el número de filas de la hoja Contenido
    Set rangoCont = wsContenido.UsedRange
    nFilasCont = rangoCont.Rows.Count
    nColumnasCont = rangoCont.Columns.Count
    
    nColumnas = nColumnas + 1
    nColumnasCont = nColumnasCont + 1
    limite = nFilas
    
    For i = 2 To limite
     Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
       'libro que ya tengo abierto
     
               dni = Cells(i, 5).Value
               
            For j = 2 To nFilasCont
              
                 'el libro que voy a abrir
               Set wsContenido = wbContenido.Worksheets("HISTORICO")
                
                 If wsContenido.Cells(j, 1).Value = dni Then
                        wsContenido.Cells(j, nColumnasCont).Value = "VERIFiCAR TIPO PROF"
                        'Regresa el control a la hoja de origen
                         Sheets("VERIF TIPOPROF").Select
                         Worksheets("VERIF TIPOPROF").Cells(i, nColumnas).Value = "ENCONTRADO-VERIF T.PROF"
                    
                  End If
                
            Next j
  
    Next i
     MsgBox "Proceso exitoso"
     Application.StatusBar = False
    Exit Sub
ControlErrorOpen:
    MsgBox "No se ha encontrado el archivo '" & contenido & "'", , "Error"
End Sub


