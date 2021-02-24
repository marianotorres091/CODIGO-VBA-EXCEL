Attribute VB_Name = "M�dulo1"
Sub SepararApeyNom()
Dim celda As Range       'celda que contiene el texto
Dim i As Integer
Dim n As Integer         'n�mero de palabras encontradas
Dim palabras() As String 'arreglo que almacenar� las palabras separadas
Dim separador As String  'separador de cada palabra
Dim texto As String      'almacena el texto a separar
Dim rango As Range
Dim cell As Range, AreaTotrim As Range

'Calculo el n�mero de filas de la hoja actual del libro que ya tengo abierto
     Set rango = ActiveSheet.UsedRange
     nFilas = rango.Rows.Count
     
     rangoTemp = "C1:C" & nFilas
     
     Set AreaTotrim = Worksheets("Hoja1").Range(rangoTemp)
     
    
    'definir el separador de palabras
    separador = "," 'espacio en blanco
    
    'Ciclo para recorrer los renglones
    For Each celda In Selection
        texto = celda.Value
        
        'Separaci�n del texto en palabras:
        palabras = Split(texto, separador)
        
        'La funci�n UBound devuelve �ndice final/mayor del arreglo
        'El �ndice en el arreglo se inicia con cero
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

