Attribute VB_Name = "Módulo1"
Sub Retiros_Ley_2871_H()
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
    Cells(2, 3).Value = "TOTALES RETIROS POR JUR"
    Cells(2, 4).Value = "TOTALES RETIROS ENTES AUTÁRQUICOS"
    
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
    Cells(16, 2).Value = "INSTITUTO DE INVEST.FORESTAL Y AGROPECUARIA"
    Cells(16, 1).Value = "15"
    Cells(17, 2).Value = "INSSEP"
    Cells(17, 1).Value = "16"
    Cells(18, 2).Value = "TRIBUNAL ELECTORAL"
    Cells(18, 1).Value = "17"
    Cells(19, 2).Value = "FISCALIA DE ESTADO"
    Cells(19, 1).Value = "18"
    Cells(20, 2).Value = "CONTADURIA GENERAL DE LA PROVINCIA"
    Cells(20, 1).Value = "19"
    Cells(21, 2).Value = "ADMINISTRACION TRIBUTARIA PROVINCIAL"
    Cells(21, 1).Value = "20"
    Cells(22, 2).Value = "POLICIA PROVINCIAL"
    Cells(22, 1).Value = "21"
    Cells(23, 2).Value = "S.U.O.P.E."
    Cells(23, 1).Value = "22"
    Cells(24, 2).Value = "MINISTERIO DE INFRAESTRUCTURA Y SERVICIOS PUBLICOS"
    Cells(24, 1).Value = "23"
    Cells(25, 2).Value = "ADMINISTRACION PROVINCIAL DEL AGUA"
    Cells(25, 1).Value = "24"
    Cells(26, 2).Value = "INSTITUTO DEL ABORIGEN CHACO"
    Cells(26, 1).Value = "25"
    Cells(27, 2).Value = "TESORERIA GENERAL"
    Cells(27, 1).Value = "26"
    Cells(28, 2).Value = "FISCALÍA DE INVESTIGACIONES ADMINISTRATIVAS"
    Cells(28, 1).Value = "27"
    Cells(29, 2).Value = "MINISTERIO DE DESARROLLO SOCIAL"
    Cells(29, 1).Value = "28"
    Cells(30, 2).Value = "MINISTERIO DE EDUCACION"
    Cells(30, 1).Value = "29"
    Cells(31, 2).Value = "EX - FONDO NACIONAL DE VIVIENDA"
    Cells(31, 1).Value = "30"
    Cells(32, 2).Value = "FONDO ESPECIAL DE RETIROS VOLUNTARIOS"
    Cells(32, 1).Value = "31"
    Cells(33, 2).Value = "MINISTERIO DE PLANIFICACION"
    Cells(33, 1).Value = "32"
    Cells(34, 2).Value = "SECRETARIA DE INVERSIONES,ASUNTOS INTERNAC.Y PROM"
    Cells(34, 1).Value = "33"
    Cells(35, 2).Value = "INSTITUTO DE CULTURA DEL CHACO"
    Cells(35, 1).Value = "34"
    Cells(36, 2).Value = "INSSEP - Pasivos"
    Cells(36, 1).Value = "35"
    Cells(37, 2).Value = "SERVICIO PENITENCIARIO PROVINCIAL"
    Cells(37, 1).Value = "36"
    Cells(38, 2).Value = "INST. PCIAL. PARA LA INCLUSION DE PERSONAS CON D"
    Cells(38, 1).Value = "37"
    Cells(39, 2).Value = "ADMINISTRACION PORTUARIA PUERTO DE BARRANQUERAS"
    Cells(39, 1).Value = "39"
    Cells(40, 2).Value = "CONSEJO DE LA MAGISTRATURA Y JURADO ENJUICIAMIENTO"
    Cells(40, 1).Value = "40"
    Cells(41, 2).Value = "SECHEEP"
    Cells(41, 1).Value = "41"
    Cells(42, 2).Value = "INSTITUTO DE TURISMO DEL CHACO"
    Cells(42, 1).Value = "42"
    Cells(43, 2).Value = "FONDO ESPECIAL RETIRO  LEY 6636                   "
    Cells(43, 1).Value = "44"
    Cells(44, 2).Value = "MINISTERIO DE DESARROLLO URBANO Y ORD. TERRITORIAL"
    Cells(44, 1).Value = "45"
    Cells(45, 2).Value = "SECRETARIA DE DERECHS HUMANOS"
    Cells(45, 1).Value = "46"
    Cells(46, 2).Value = "INSTITUTO DEL DEFENSOR DEL PUEBLO"
    Cells(46, 1).Value = "47"
    Cells(47, 2).Value = "INSTITUTO DEL DEPORTE"
    Cells(47, 1).Value = "48"
    Cells(48, 2).Value = "ESCUELA DE GOBIERNO DE LA PROVICNIA DEL CHACO"
    Cells(48, 1).Value = "49"
    Cells(49, 2).Value = "MEC. PROV. DE PREV. DE LA TORTURA Y OTROS TRATOS Y PENAS CRUELES INHUMANOS Y/O DEGRADANTES"
    Cells(49, 1).Value = "50"
    Cells(50, 2).Value = "INST. PROVINCIAL DE ADMIN. PUBLICA DEL CHACO"
    Cells(50, 1).Value = "51"
    Cells(51, 2).Value = "INST. DESARROLLO RURAL Y AGRICULTURA  FLIAR"
    Cells(51, 1).Value = "52"
    Cells(52, 2).Value = "SECRETARIA DE MUNICIPIOS Y CIUDADES"
    Cells(52, 1).Value = "53"
    Cells(53, 2).Value = "SECRETARIA DE EMPLEO Y TRABAJO"
    Cells(53, 1).Value = "54"
    Cells(54, 2).Value = "PUERTO LAS PALMAS en creacion (D. 1094/18 - D. 2963/19, depende de la 23)"
    Cells(54, 1).Value = "55"
    Cells(55, 2).Value = "TOTALES"
    
   'PUESTA A CERO ENTES AUTARTICOS
   'JUR10
   Sheets("REPORTE RETIROS").Cells(11, 4).Value = 0
   'JUR12
   Sheets("REPORTE RETIROS").Cells(13, 4).Value = 0
   'JUR 13
   Sheets("REPORTE RETIROS").Cells(14, 4).Value = 0
   'JUR 14
   Sheets("REPORTE RETIROS").Cells(15, 4).Value = 0
   'JUR 15
   Sheets("REPORTE RETIROS").Cells(16, 4).Value = 0
   'JUR 16
   Sheets("REPORTE RETIROS").Cells(17, 4).Value = 0
   'JUR 20
   Sheets("REPORTE RETIROS").Cells(21, 4).Value = 0
   'JUR 21
   Sheets("REPORTE RETIROS").Cells(22, 4).Value = 0
   'JUR 24
   Sheets("REPORTE RETIROS").Cells(25, 4).Value = 0
   'JUR 34
   Sheets("REPORTE RETIROS").Cells(35, 4).Value = 0
   'JUR 35
   Sheets("REPORTE RETIROS").Cells(36, 4).Value = 0
   'JUR 36
   Sheets("REPORTE RETIROS").Cells(37, 4).Value = 0
   'JUR 37
   Sheets("REPORTE RETIROS").Cells(38, 4).Value = 0
   'JUR 39
   Sheets("REPORTE RETIROS").Cells(39, 4).Value = 0
   'JUR 42
   Sheets("REPORTE RETIROS").Cells(42, 4).Value = 0
   'JUR 48
   Sheets("REPORTE RETIROS").Cells(47, 4).Value = 0
   'JUR 49
   Sheets("REPORTE RETIROS").Cells(48, 4).Value = 0
   'JUR 51
   Sheets("REPORTE RETIROS").Cells(50, 4).Value = 0
   'JUR 52
   Sheets("REPORTE RETIROS").Cells(51, 4).Value = 0
   
    limite = 54
    For i = 3 To limite
     Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
     
     Sheets("REPORTE RETIROS").Cells(i, 3).Value = 0
     
       For j = 2 To nFilas
       
          If Sheets("ptatipo47").Cells(j, 2).Value = Sheets("REPORTE RETIROS").Cells(i, 1).Value Then
      
            Sheets("REPORTE RETIROS").Cells(i, 3).Value = Sheets("REPORTE RETIROS").Cells(i, 3).Value + 1
            
            reporte = Sheets("REPORTE RETIROS").Cells(i, 1).Value
            
            If reporte = 10 Or reporte = 12 Or reporte = 13 Or reporte = 14 Or reporte = 15 Or reporte = 16 Or reporte = 20 Or reporte = 21 Or reporte = 24 Or reporte = 34 Or reporte = 35 Or reporte = 36 Or reporte = 37 Or reporte = 39 Or reporte = 42 Or reporte = 48 Or reporte = 49 Or reporte = 51 Or reporte = 52 Then
             Sheets("REPORTE RETIROS").Cells(i, 4).Value = Sheets("REPORTE RETIROS").Cells(i, 4).Value + 1
            End If
            
          End If
          
        Next j
     Sheets("REPORTE RETIROS").Cells(55, 3).Value = Sheets("REPORTE RETIROS").Cells(55, 3).Value + Sheets("REPORTE RETIROS").Cells(i, 3).Value
     Sheets("REPORTE RETIROS").Cells(55, 4).Value = Sheets("REPORTE RETIROS").Cells(55, 4).Value + Sheets("REPORTE RETIROS").Cells(i, 4).Value
    Next i
       
             
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    Application.StatusBar = False
    Exit Sub

End Sub



