Attribute VB_Name = "Módulo1"
Sub DNI_iguales_TipoProf_dif2()
    Dim nFilas As Long
    Dim nColumnas As Long
    'Dim i As Integer
    Dim rango As Range
    Dim band As Double
    
    
    'ATECION!!! PRIMERO APLICAR EL ORDENAMIENTO POR DNI Y UNA VEZ QUE APLIQUE VUELVO A APLICAR EL ORDENAMIENTO POR TIPOPROF!!!
    'SI APLICO EL ORDENAMIENTO POR DNI Y TIPOPROF JUNTOS NO FUNCIONA EL CODIGO
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
    Sheets.Add.Name = "ERRORES"
    
    filaResultado = 1
    
    Cells(filaResultado, 1).Value = "CUOF"
    Cells(filaResultado, 2).Value = "ANEXO"
    Cells(filaResultado, 3).Value = "AÑO"
    Cells(filaResultado, 4).Value = "MES"
    Cells(filaResultado, 5).Value = "DNI"
    Cells(filaResultado, 6).Value = "APELLIDO Y NOMBRE"
    
    Sheets("GUARDIAS 04-20").Select
    
    band = False
    limite = (nFilas - 1)
    For i = 2 To limite
     Application.StatusBar = Format(i / limite, "0.0%") & " " & "Completo"
     
    
     If Cells(i, nColumnas + 1).Value <> "DNI= - TIPOPROF DIST" Then
        
        cuof = Cells(i, 1).Value
        Doc = Cells(i, 5).Value
        año = Cells(i, 3).Value
        mes = Cells(i, 4).Value
        anexo = Cells(i, 2).Value
        nombre = Cells(i, 6).Value
        tipoProf = Cells(i, 7).Value
        
        For j = (i + 1) To nFilas
            If Doc = Cells(j, 5).Value And tipoProf <> Cells(j, 7).Value Then
             If tipoProf = "A" Then
               If Cells(j, 7).Value <> "D" Then
                    Sheets("ERRORES").Select
                     'Calcular el número de filas de la hoja ERRORES
                        Set rango = ActiveSheet.UsedRange
                        nFilasE = rango.Rows.Count
                        nColumnasE = rango.Columns.Count
                        
                     For e = 1 To nFilasE
                      If Cells(e, 5).Value = Doc Then
                       band = True
                      End If
                     Next e
                     
                     If band = False Then
                        filaResultado = filaResultado + 1
                        Cells(filaResultado, 1).Value = cuof
                        Cells(filaResultado, 2).Value = anexo
                        Cells(filaResultado, 3).Value = año
                        Cells(filaResultado, 4).Value = mes
                        Cells(filaResultado, 5).Value = Doc
                        Cells(filaResultado, 6).Value = nombre
                      Else
                      band = False
                    End If
                    
                    Sheets("GUARDIAS 04-20").Select
                    
                    Cells(j, 5).Interior.Color = RGB(240, 243, 121)
                    Cells(i, 5).Interior.Color = RGB(240, 243, 121)
                    Cells(j, nColumnas + 1).Value = "DNI= - TIPOPROF DIST"
                    If Cells(j, nColumnas + 2).Value = "" Then
                        Cells(j, nColumnas + 2).Value = i
                        Cells(i, nColumnas + 2).Value = i
                       Else
                        Cells(i, nColumnas + 2).Value = Cells(j, nColumnas + 2).Value
                     End If
                    Cells(i, nColumnas + 1).Value = "DNI= - TIPOPROF DIST"
               End If
               Else
                    Sheets("ERRORES").Select
                     'Calcular el número de filas de la hoja ERRORES
                        Set rango = ActiveSheet.UsedRange
                        nFilasE = rango.Rows.Count
                        nColumnasE = rango.Columns.Count
                        
                     For e = 1 To nFilasE
                      If Cells(e, 5).Value = Doc Then
                       band = True
                      End If
                     Next e
                     
                     If band = False Then
                        filaResultado = filaResultado + 1
                        Cells(filaResultado, 1).Value = cuof
                        Cells(filaResultado, 2).Value = anexo
                        Cells(filaResultado, 3).Value = año
                        Cells(filaResultado, 4).Value = mes
                        Cells(filaResultado, 5).Value = Doc
                        Cells(filaResultado, 6).Value = nombre
                      Else
                      band = False
                    End If
                    
                    Sheets("GUARDIAS 04-20").Select
                    
                    Cells(j, 5).Interior.Color = RGB(240, 243, 121)
                    Cells(i, 5).Interior.Color = RGB(240, 243, 121)
                    Cells(j, nColumnas + 1).Value = "DNI= - TIPOPROF DIST"
                     If Cells(j, nColumnas + 2).Value = "" Then
                        Cells(j, nColumnas + 2).Value = i
                        Cells(i, nColumnas + 2).Value = i
                       Else
                        Cells(i, nColumnas + 2).Value = Cells(j, nColumnas + 2).Value
                     End If
                    Cells(i, nColumnas + 1).Value = "DNI= - TIPOPROF DIST"
             End If
            End If
        Next j
      End If
  
    Next i
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Application.StatusBar = False
End Sub
