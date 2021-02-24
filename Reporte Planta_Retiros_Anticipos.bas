Attribute VB_Name = "Módulo1"
    Sub Reporte_Planta_Retiros_Anticipos()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim tempFecha As Date
    Dim i, RowNumber As Long
    Dim filaResultado As Long
    Dim filaCopia As Long
    Dim columnaCopia As Long
    Dim wbContenido As Worksheet
       

    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    

     Sheets.Add.Name = "REPORTE RETIROS"
 
    
    Cells(2, 1).Value = "JUR"
    Cells(2, 2).Value = "DENOMINACIÓN"
    Cells(2, 3).Value = "PLANTA PERMANENTE"
    Cells(2, 4).Value = "LEY DE RETIRO 3852"
    Cells(2, 5).Value = "LEY DE RETIRO 4256"
    Cells(2, 6).Value = "LEY DE RETIRO 6635"
    Cells(2, 7).Value = "RETIRO LEY 2871-H"
    Cells(2, 8).Value = "ANTICIPO PREVISIONAL"
    
    Cells(3, 2).Value = "SECRETARIA GENERAL DE GOBIERNO Y COORDINACION"
    Cells(3, 1).Value = "2"
    Cells(4, 2).Value = "MIN. DE GOBIERNO, JUSTICIA Y RELACIÓN CON LA COMUNIDAD"
    Cells(4, 1).Value = "3"
    Cells(5, 2).Value = "MINISTERIO PLANIFICACION Y ECONOMIA"
    Cells(5, 1).Value = "4"
    Cells(6, 2).Value = "MINISTERIO DE PRODUCCION"
    Cells(6, 1).Value = "5"
    Cells(7, 2).Value = "MINISTERIO DE SALUD PUBLICA"
    Cells(7, 1).Value = "6"
    Cells(8, 2).Value = "MINISTRIO SEGURIDAD Y JUSTICIA"
    Cells(8, 1).Value = "7"
    Cells(9, 2).Value = "TRIBUNAL DE CUENTAS"
    Cells(9, 1).Value = "8"
    Cells(10, 2).Value = "PODER JUDICIAL"
    Cells(10, 1).Value = "9"
    Cells(11, 2).Value = "INSTITUTO PCIAL. DE DESARROLLO URB. Y VIVIENDA"
    Cells(11, 1).Value = "10"
    Cells(12, 2).Value = "MINISTERIO DE INDUSTRIA, COMERCIO Y SERVICIOS"
    Cells(12, 1).Value = "11"
    Cells(13, 2).Value = "LOTERIA CHAQUEÑA"
    Cells(13, 1).Value = "12"
    Cells(14, 2).Value = "DIRECCION DE VIALIDAD PROVINCIAL"
    Cells(14, 1).Value = "13"
    Cells(15, 2).Value = "INSTITUTO DE COLONIZACION"
    Cells(15, 1).Value = "14"
    Cells(16, 2).Value = "INSTITUTO DE INVEST.FORESTAL Y AGROPECUA"
    Cells(16, 1).Value = "15"
    Cells(17, 2).Value = "INSSEP"
    Cells(17, 1).Value = "16"
    Cells(18, 2).Value = "TRIBUNAL ELECTORAL"
    Cells(18, 1).Value = "17"
    Cells(19, 2).Value = "FISCALIA DE ESTADO"
    Cells(19, 1).Value = "18"
    Cells(20, 2).Value = "CONTADURIA GENERAL"
    Cells(20, 1).Value = "19"
    Cells(21, 2).Value = "ADMINISTRACION TRIBUTARIA PROVINCIAL"
    Cells(21, 1).Value = "20"
    Cells(22, 2).Value = "POLICIA PROVINCIAL"
    Cells(22, 1).Value = "21"
    Cells(23, 2).Value = "MINISTERIO DE INFRAESTRUCTURA Y SERVICIOS PUBLICOS"
    Cells(23, 1).Value = "23"
    Cells(24, 2).Value = "ADMINISTRACION PROVINCIAL DEL AGUA"
    Cells(24, 1).Value = "24"
    Cells(25, 2).Value = "INSTITUTO DEL ABORIGEN CHACO"
    Cells(25, 1).Value = "25"
    Cells(26, 2).Value = "TESORERIA GENERAL"
    Cells(26, 1).Value = "26"
    Cells(27, 2).Value = "TESORERIA GENERAL"
    Cells(27, 1).Value = "27"
    Cells(28, 2).Value = "MINISTERIO DE DESARROLLO SOCIAL"
    Cells(28, 1).Value = "28"
    Cells(29, 2).Value = "MINISTERIO DE EDUCACION"
    Cells(29, 1).Value = "29"
    Cells(30, 2).Value = "FONDO ESPECIAL DE RETIROS VOLUNTARIOS"
    Cells(30, 1).Value = "31"
    Cells(31, 2).Value = "MINISTERIO DE PLANIFICACION"
    Cells(31, 1).Value = "32"
    Cells(32, 2).Value = "SECRETARIA DE INVERSIONES,ASUNTOS INTERNAC.Y PROM"
    Cells(32, 1).Value = "33"
    Cells(33, 2).Value = "INSTITUTO DE CULTURA DEL CHACO"
    Cells(33, 1).Value = "34"
    Cells(34, 2).Value = "SERVICIO PENITENCIARIO PROVINCIAL"
    Cells(34, 1).Value = "36"
    Cells(35, 2).Value = "INST. PCIAL. PARA LA INCLUSION DE PERSONAS CON D"
    Cells(35, 1).Value = "37"
    Cells(36, 2).Value = "ADMINISTRACION PORTUARIA PUERTO DE BARRANQUERAS"
    Cells(36, 1).Value = "39"
    Cells(37, 2).Value = "CONSEJO DE LA MAGISTRATURA Y JURADO ENJUICIAMIENTO"
    Cells(37, 1).Value = "40"
    Cells(38, 2).Value = "INSTITUTO DE TURISMO DEL CHACO"
    Cells(38, 1).Value = "42"
    Cells(39, 2).Value = "FONDO ESPECIAL RETIRO  LEY 6636                   "
    Cells(39, 1).Value = "44"
    Cells(40, 2).Value = "MINISTERIO DE DESARROLLO URBANO Y ORD. TERRITORIAL"
    Cells(40, 1).Value = "45"
    Cells(41, 2).Value = "SECRETARIA DE DERECHS HUMANOS"
    Cells(41, 1).Value = "46"
    Cells(42, 2).Value = "INSTITUTO DEL DEFENSOR DEL PUEBLO"
    Cells(42, 1).Value = "47"
    Cells(43, 2).Value = "INSTITUTO DEL DEPORTE"
    Cells(43, 1).Value = "48"
    Cells(44, 2).Value = "ESCUELA DE GOBIERNO DE LA PROVICNIA DEL CHACO"
    Cells(44, 1).Value = "49"
    Cells(45, 2).Value = "MEC. PROV. DE PREV. DE LA TORTURA Y OTROS TRATOS"
    Cells(45, 1).Value = "50"
    Cells(46, 2).Value = "INST. PROVINCIAL DE ADMIN. PUBLICA DEL CHACO"
    Cells(46, 1).Value = "51"
    Cells(47, 2).Value = "INST. DESARROLLO RURAL Y AGRICULTURA  FLIAR"
    Cells(47, 1).Value = "52"
    Cells(48, 2).Value = "SECRETARIA DE MUNICIPIOS Y CIUDADES"
    Cells(48, 1).Value = "53"
    Cells(49, 2).Value = "SECRETARIA DE EMPLEO Y TRABAJO"
    Cells(49, 1).Value = "54"
    Cells(50, 2).Value = "PUERTO LAS PALMAS en creacion (D. 1094/18 - D. 2963/19, depende de la 23)"
    Cells(50, 1).Value = "55"
    Cells(51, 2).Value = "TOTALES"
    
   
   
    limite = 50
    For i = 3 To limite
     Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
     Sheets("REPORTE RETIROS").Cells(i, 3).Value = 0
     Sheets("REPORTE RETIROS").Cells(i, 4).Value = 0
     Sheets("REPORTE RETIROS").Cells(i, 5).Value = 0
     Sheets("REPORTE RETIROS").Cells(i, 6).Value = 0
     Sheets("REPORTE RETIROS").Cells(i, 7).Value = 0
     Sheets("REPORTE RETIROS").Cells(i, 8).Value = 0
       For j = 2 To 40387
        If Sheets("totales").Cells(j, 7).Value = 1 Then
          If Sheets("totales").Cells(j, 2).Value = Sheets("REPORTE RETIROS").Cells(i, 1).Value Then
      
            Sheets("REPORTE RETIROS").Cells(i, 3).Value = Sheets("REPORTE RETIROS").Cells(i, 3).Value + 1
            
          End If
         Else
         If Sheets("totales").Cells(j, 7).Value = 3 Then
          If Sheets("totales").Cells(j, 2).Value = Sheets("REPORTE RETIROS").Cells(i, 1).Value Then
           Sheets("REPORTE RETIROS").Cells(i, 4).Value = Sheets("REPORTE RETIROS").Cells(i, 4).Value + 1
          End If
          
          Else
          
           If Sheets("totales").Cells(j, 7).Value = 16 Then
            If Sheets("totales").Cells(j, 2).Value = Sheets("REPORTE RETIROS").Cells(i, 1).Value Then
               Sheets("REPORTE RETIROS").Cells(i, 5).Value = Sheets("REPORTE RETIROS").Cells(i, 5).Value + 1
            End If
            
            Else
            
             If Sheets("totales").Cells(j, 7).Value = 41 Then
              If Sheets("totales").Cells(j, 2).Value = Sheets("REPORTE RETIROS").Cells(i, 1).Value Then
               Sheets("REPORTE RETIROS").Cells(i, 6).Value = Sheets("REPORTE RETIROS").Cells(i, 6).Value + 1
              End If
              
              Else
              
               If Sheets("totales").Cells(j, 7).Value = 47 Then
                If Sheets("totales").Cells(j, 2).Value = Sheets("REPORTE RETIROS").Cells(i, 1).Value Then
                 Sheets("REPORTE RETIROS").Cells(i, 7).Value = Sheets("REPORTE RETIROS").Cells(i, 7).Value + 1
                End If
               
               Else
                 If Sheets("totales").Cells(j, 7).Value = 2 Then
                  If Sheets("totales").Cells(j, 2).Value = Sheets("REPORTE RETIROS").Cells(i, 1).Value Then
                   Sheets("REPORTE RETIROS").Cells(i, 8).Value = Sheets("REPORTE RETIROS").Cells(i, 8).Value + 1
                  End If
                 End If
              End If
             End If
           End If
         End If
        End If
        Next j
     Sheets("REPORTE RETIROS").Cells(51, 3).Value = Sheets("REPORTE RETIROS").Cells(51, 3).Value + Sheets("REPORTE RETIROS").Cells(i, 3).Value
     Sheets("REPORTE RETIROS").Cells(51, 4).Value = Sheets("REPORTE RETIROS").Cells(51, 4).Value + Sheets("REPORTE RETIROS").Cells(i, 4).Value
     Sheets("REPORTE RETIROS").Cells(51, 5).Value = Sheets("REPORTE RETIROS").Cells(51, 5).Value + Sheets("REPORTE RETIROS").Cells(i, 5).Value
     Sheets("REPORTE RETIROS").Cells(51, 6).Value = Sheets("REPORTE RETIROS").Cells(51, 6).Value + Sheets("REPORTE RETIROS").Cells(i, 6).Value
     Sheets("REPORTE RETIROS").Cells(51, 7).Value = Sheets("REPORTE RETIROS").Cells(51, 7).Value + Sheets("REPORTE RETIROS").Cells(i, 7).Value
     Sheets("REPORTE RETIROS").Cells(51, 8).Value = Sheets("REPORTE RETIROS").Cells(51, 8).Value + Sheets("REPORTE RETIROS").Cells(i, 8).Value
    Next i
       
             
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Application.StatusBar = False
    Exit Sub

End Sub

