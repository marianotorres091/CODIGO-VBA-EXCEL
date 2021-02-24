Attribute VB_Name = "Módulo1"
Sub Duplicados_Columna()
Dim fila As Long
Dim final As Long
    
final = Range("A1").End(xlDown).Row
 For fila = 1 To final
    If Application.WorksheetFunction.CountIf(Range("A1:A" & final), Range("A" & fila)) > 1 Then
        Range("A" & fila).Interior.Color = RGB(200, 200, 200)
    End If
 Next fila
 End Sub
