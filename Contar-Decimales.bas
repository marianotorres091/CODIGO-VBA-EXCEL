Attribute VB_Name = "Módulo1"
Sub Contar_decimales()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim i, cuoc, rj, unidad, importe, vto As Long
   
   'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    
 limite = nFilas
 
 For j = 2 To limite
  Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
    dec = False
    palabra = Cells(j, 12).Value
    cont = 0
    tamaño = Len(palabra)
    For i = 1 To tamaño
     Caracter = Mid(CStr(palabra), i, 1)
        If Caracter = "," Then
          dec = True
        End If
        
        If dec Then
         If Caracter <> "," Then
         cont = cont + 1
         End If
        End If
             
    Next i
    
    If cont > 2 Then
    Cells(j, nColumnas + 1) = cont
    End If
  Next j
     MsgBox "Proceso exitosa"
      Application.StatusBar = False
End Sub

    
