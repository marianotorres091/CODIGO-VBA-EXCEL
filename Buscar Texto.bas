Attribute VB_Name = "Módulo1"
Sub BUSCAR_TEXTO()
Dim busqueda As String
busqueda = Application.InputBox(Prompt:="¿Que estas buscando?", Title:="Busqueda")

With ActiveSheet.Range("A1:XFD1048576")

    Set c = .Find(What:=busqueda, _
                   LookIn:=xlValues, _
                   LookAt:=xlPart, _
                   SearchDirection:=xlNext, _
                   SearchFormat:=False)
    If Not c Is Nothing Then
     primerDireccion = c.Address
     Do
        Range(c.Address).Select
        Selection.Interior.Color = RGB(0, 255, 0)
        Set c = .FindNext(c)
        If c Is Nothing Then
           GoTo finBusqueda
        End If
     Loop While Not c Is Nothing And c.Address <> primerDireccion
    Else
     MsgBox "No se encontraron resultados"
   End If
finBusqueda:
End With
End Sub
