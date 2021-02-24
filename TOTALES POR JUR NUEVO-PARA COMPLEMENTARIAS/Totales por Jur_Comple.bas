Attribute VB_Name = "Módulo1"
Sub Total_por_Jur()
    Dim rango As Range
    Dim nFilas As Long
    Dim nColumnas As Long
    Dim tempFecha As Date
    Dim i, RowNumber As Long
    Dim filaResultado As Long
    Dim filaCopia As Long
    Dim columnaCopia As Long
    Dim wbContenido As Worksheet
       
    'PRIMERO EJECUTAR ESTE CODIGO ANTES DE EJECUTAR EL Total_por_Jur_Auditoria_de_Liq, EL Total_por_Jur_Seguridad_Sistemas y Total_por_Jur_No_Tratadas.
    'DESPUES EJECUTAR EL Total_por_Jur_Auditoria_de_Liq, EL Total_por_Jur_Seguridad_Sistemas y Total_por_Jur_No_Tratadas.
    'POR ULTIMO DARLE EL FORMATO AL REPORTE
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    nColumnas = rango.Columns.Count
    

     Sheets.Add.Name = "REPORTE"
 
    
    Cells(2, 1).Value = "JUR"
    Cells(2, 2).Value = "DENOMINACIÓN"
    Cells(2, 3).Value = "LIQUIDADO"
    Cells(3, 2).Value = "SECRETARIA GENERAL DE GOBIERNO Y COORDINACION"
    Cells(3, 1).Value = "2"
    Cells(4, 2).Value = "MINISTERIO DE GOBIERNO"
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
    Cells(14, 2).Value = "VIALIDAD PROVINCIAL"
    Cells(14, 1).Value = "13"
    Cells(15, 2).Value = "INSTITUTO DE COLONIZACION"
    Cells(15, 1).Value = "14"
    Cells(16, 2).Value = "INSTITUTO DE INVEST.FORESTAL Y AGROPECUARIA"
    Cells(16, 1).Value = "15"
    Cells(17, 2).Value = "TRIBUNAL ELECTORAL"
    Cells(17, 1).Value = "17"
    Cells(18, 2).Value = "FISCALIA DE ESTADO"
    Cells(18, 1).Value = "18"
    Cells(19, 2).Value = "CONTADURIA GENERAL"
    Cells(19, 1).Value = "19"
    Cells(20, 2).Value = "ADMINISTRACION TRIBUTARIA PROVINCIAL"
    Cells(20, 1).Value = "20"
    Cells(21, 2).Value = "POLICIA PROVINCIAL"
    Cells(21, 1).Value = "21"
    Cells(22, 2).Value = "MINISTERIO DE INFRAESTRUCTURA Y SERVICIOS PUBLICOS"
    Cells(22, 1).Value = "23"
    Cells(23, 2).Value = "ADMINISTRACION PROVINCIAL DEL AGUA"
    Cells(23, 1).Value = "24"
    Cells(24, 2).Value = "INSTITUTO DEL ABORIGEN CHACO"
    Cells(24, 1).Value = "25"
    Cells(25, 2).Value = "TESORERIA GENERAL"
    Cells(25, 1).Value = "26"
    Cells(26, 2).Value = "FISCALIA DE INVESTIGACIONES ADMINISTRATIVAS"
    Cells(26, 1).Value = "27"
    Cells(27, 2).Value = "MINISTERIO DE DESARROLLO SOCIAL"
    Cells(27, 1).Value = "28"
    Cells(28, 2).Value = "MINISTERIO DE EDUCACION"
    Cells(28, 1).Value = "29"
    Cells(29, 2).Value = "FONDO ESPECIAL DE RETIROS VOLUNTARIOS"
    Cells(29, 1).Value = "31"
    Cells(30, 2).Value = "MINISTERIO DE PLANIFICACION"
    Cells(30, 1).Value = "32"
    Cells(31, 2).Value = "SECRETARIA DE INVERSIONES,ASUNTOS INTERNAC.Y PROM"
    Cells(31, 1).Value = "33"
    Cells(32, 2).Value = "INSTITUTO DE CULTURA DEL CHACO"
    Cells(32, 1).Value = "34"
    Cells(33, 2).Value = "SERVICIO PENITENCIARIO PROVINCIAL"
    Cells(33, 1).Value = "36"
    Cells(34, 2).Value = "INST. PCIAL. PARA LA INCLUSION DE PERSONAS CON D"
    Cells(34, 1).Value = "37"
    Cells(35, 2).Value = "ADMINISTRACION PORTUARIA PUERTO DE BARRANQUERAS"
    Cells(35, 1).Value = "39"
    Cells(36, 2).Value = "CONSEJO DE LA MAGISTRATURA Y JURADO ENJUICIAMIENTO"
    Cells(36, 1).Value = "40"
    Cells(37, 2).Value = "INSTITUTO DE TURISMO DEL CHACO"
    Cells(37, 1).Value = "42"
    Cells(38, 2).Value = "FONDO ESPECIAL RETIRO  LEY 6636                   "
    Cells(38, 1).Value = "44"
    Cells(39, 2).Value = "MINISTERIO DE DESARROLLO URBANO Y ORD. TERRITORIAL"
    Cells(39, 1).Value = "45"
    Cells(40, 2).Value = "SECRETARIA DE DERECHS HUMANOS"
    Cells(40, 1).Value = "46"
    Cells(41, 2).Value = "INSTITUTO DEL DEFENSOR DEL PUEBLO"
    Cells(41, 1).Value = "47"
    Cells(42, 2).Value = "INSTITUTO DEL DEPORTE"
    Cells(42, 1).Value = "48"
    Cells(43, 2).Value = "ESCUELA DE GOBIERNO DE LA PROVICNIA DEL CHACO"
    Cells(43, 1).Value = "49"
    Cells(44, 2).Value = "MEC. PROV. DE PREV. DE LA TORTURA Y OTROS TRATOS"
    Cells(44, 1).Value = "50"
    Cells(45, 2).Value = "INST. PROVINCIAL DE ADMIN. PUBLICA DEL CHACO"
    Cells(45, 1).Value = "51"
    Cells(46, 2).Value = "INST. DESARROLLO RURAL Y AGRICULTURA  FLIAR"
    Cells(46, 1).Value = "52"
    Cells(47, 2).Value = "SECRETARIA DE MUNICIPIOS Y CIUDADES"
    Cells(47, 1).Value = "53"
    Cells(48, 2).Value = "SECRETARIA DE EMPLEO Y TRABAJO"
    Cells(48, 1).Value = "54"
    Cells(49, 2).Value = "PUERTO LAS PALMAS en creacion (D. 1094/18 - D. 2963/19, depende de la 23)"
    Cells(49, 1).Value = "55"
    
    limite = 49
    
    For i = 3 To limite
    
      Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
      
      Sheets("REPORTE").Cells(i, 3).Value = 0
      
       For j = 2 To nFilas
       
          If Sheets("Hoja1").Cells(j, 2).Value = Sheets("REPORTE").Cells(i, 1).Value Then
            
            If Sheets("Hoja1").Cells(j, 10).Value = 1 Or Sheets("Hoja1").Cells(j, 10).Value = 0 Then
          
              Sheets("REPORTE").Cells(i, 3).Value = Sheets("REPORTE").Cells(i, 3).Value + Sheets("Hoja1").Cells(j, 12).Value
             Else
              Sheets("REPORTE").Cells(i, 3).Value = Sheets("REPORTE").Cells(i, 3).Value - Sheets("Hoja1").Cells(j, 12).Value
            End If
          End If
          
        Next j
        
       Sheets("REPORTE").Cells(50, 3).Value = Sheets("REPORTE").Cells(50, 3).Value + Sheets("REPORTE").Cells(i, 3).Value
        
    Next i
       
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    
    Application.StatusBar = False
    
    Exit Sub

End Sub

Sub Total_por_Jur_Auditoria_de_Liq()

    Cells(2, 6).Value = "DPTO. AUDITORIA DE LIQUIDACIONES"
    Cells(3, 6).Value = "JUR"
    Cells(3, 7).Value = "DENOMINACIÓN"
    Cells(3, 8).Value = "LIQUIDADO"
    
    Cells(4, 7).Value = "SECRETARIA GENERAL DE GOBIERNO Y COORDINACION"
    Cells(4, 6).Value = "2"
    Cells(4, 8).Value = Cells(3, 3).Value
    Cells(5, 7).Value = "MINISTERIO PLANIFICACION Y ECONOMIA"
    Cells(5, 6).Value = "4"
    Cells(5, 8).Value = Cells(5, 3).Value
    Cells(6, 7).Value = "MINISTERIO DE SALUD PUBLICA"
    Cells(6, 6).Value = "6"
    Cells(6, 8).Value = Cells(7, 3).Value
    Cells(7, 7).Value = "MINISTRIO SEGURIDAD Y JUSTICIA"
    Cells(7, 6).Value = "7"
    Cells(7, 8).Value = Cells(8, 3).Value
    Cells(8, 7).Value = "INSTITUTO PCIAL. DE DESARROLLO URB. Y VIVIENDA"
    Cells(8, 6).Value = "10"
    Cells(8, 8).Value = Cells(11, 3).Value
    Cells(9, 7).Value = "MINISTERIO DE INDUSTRIA, COMERCIO Y SERVICIOS"
    Cells(9, 6).Value = "11"
    Cells(9, 8).Value = Cells(12, 3).Value
    Cells(10, 7).Value = "LOTERIA CHAQUEÑA"
    Cells(10, 6).Value = "12"
    Cells(10, 8).Value = Cells(13, 3).Value
    Cells(11, 7).Value = "POLICIA PROVINCIAL"
    Cells(11, 6).Value = "21"
    Cells(11, 8).Value = Cells(21, 3).Value
    Cells(12, 7).Value = "SERVICIO PENITENCIARIO PROVINCIAL"
    Cells(12, 6).Value = "36"
    Cells(12, 8).Value = Cells(33, 3).Value
    Cells(13, 7).Value = "SECRETARIA DE DERECHS HUMANOS"
    Cells(13, 6).Value = "46"
    Cells(13, 8).Value = Cells(40, 3).Value
    Cells(14, 7).Value = "ESCUELA DE GOBIERNO DE LA PROVICNIA DEL CHACO"
    Cells(14, 6).Value = "49"
    Cells(14, 8).Value = Cells(43, 3).Value
    Cells(15, 7).Value = "INST. PROVINCIAL DE ADMIN. PUBLICA DEL CHACO"
    Cells(15, 6).Value = "51"
    Cells(15, 8).Value = Cells(45, 3).Value
    Cells(16, 7).Value = "SECRETARIA DE EMPLEO Y TRABAJO"
    Cells(16, 6).Value = "54"
    Cells(16, 8).Value = Cells(48, 3).Value
    Cells(17, 6).Value = "TOTAL"
    
    Cells(17, 8).Value = 0
    
    For i = 4 To 16
     Cells(17, 8).Value = Cells(17, 8).Value + Cells(i, 8).Value
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    
End Sub
    
Sub Total_por_Jur_Seguridad_Sistemas()
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
    
 
    Cells(21, 6).Value = "DPTO. SEGURIDAD DE SISTEMAS"
    Cells(22, 6).Value = "JUR"
    Cells(22, 7).Value = "DENOMINACIÓN"
    Cells(22, 8).Value = "LIQUIDADO"
    
    Cells(23, 7).Value = "MINISTERIO DE GOBIERNO"
    Cells(23, 6).Value = "3"
    Cells(23, 8).Value = Cells(4, 3).Value
    Cells(24, 7).Value = "MINISTERIO DE PRODUCCION"
    Cells(24, 6).Value = "5"
    Cells(24, 8).Value = Cells(6, 3).Value
    Cells(25, 7).Value = "VIALIDAD PROVINCIAL"
    Cells(25, 6).Value = "13"
    Cells(25, 8).Value = Cells(14, 3).Value
    Cells(26, 7).Value = "INSTITUTO DE COLONIZACION"
    Cells(26, 6).Value = "14"
    Cells(26, 8).Value = Cells(15, 3).Value
    Cells(27, 7).Value = "INSTITUTO DE INVEST.FORESTAL Y AGROPECUARIA"
    Cells(27, 6).Value = "15"
    Cells(27, 8).Value = Cells(16, 3).Value
    Cells(28, 7).Value = "CONTADURIA GENERAL"
    Cells(28, 6).Value = "19"
    Cells(28, 8).Value = Cells(19, 3).Value
    Cells(29, 7).Value = "ADMINISTRACION TRIBUTARIA PROVINCIAL"
    Cells(29, 6).Value = "20"
    Cells(29, 8).Value = Cells(20, 3).Value
    Cells(30, 7).Value = "MINISTERIO DE INFRAESTRUCTURA Y SERVICIOS PUBLICOS"
    Cells(30, 6).Value = "23"
    Cells(30, 8).Value = Cells(22, 3).Value
    Cells(31, 7).Value = "ADMINISTRACION PROVINCIAL DEL AGUA"
    Cells(31, 6).Value = "24"
    Cells(31, 8).Value = Cells(23, 3).Value
    Cells(32, 7).Value = "INSTITUTO DEL ABORIGEN CHACO"
    Cells(32, 6).Value = "25"
    Cells(32, 8).Value = Cells(24, 3).Value
    Cells(33, 7).Value = "TESORERIA GENERAL"
    Cells(33, 6).Value = "26"
    Cells(33, 8).Value = Cells(25, 3).Value
    Cells(34, 7).Value = "MINISTERIO DE DESARROLLO SOCIAL"
    Cells(34, 6).Value = "28"
    Cells(34, 8).Value = Cells(27, 3).Value
    Cells(35, 7).Value = "MINISTERIO DE EDUCACION"
    Cells(35, 6).Value = "29"
    Cells(35, 8).Value = Cells(28, 3).Value
    Cells(36, 7).Value = "MINISTERIO DE PLANIFICACION"
    Cells(36, 6).Value = "32"
    Cells(36, 8).Value = Cells(30, 3).Value
    Cells(37, 7).Value = "SECRETARIA DE INVERSIONES,ASUNTOS INTERNAC.Y PROM"
    Cells(37, 6).Value = "33"
    Cells(37, 8).Value = Cells(31, 3).Value
    Cells(38, 7).Value = "INSTITUTO DE CULTURA DEL CHACO"
    Cells(38, 6).Value = "34"
    Cells(38, 8).Value = Cells(32, 3).Value
    Cells(39, 7).Value = "INST. PCIAL. PARA LA INCLUSION DE PERSONAS CON D"
    Cells(39, 6).Value = "37"
    Cells(39, 8).Value = Cells(34, 3).Value
    Cells(40, 7).Value = "ADMINISTRACION PORTUARIA PUERTO DE BARRANQUERAS"
    Cells(40, 6).Value = "39"
    Cells(40, 8).Value = Cells(35, 3).Value
    Cells(41, 7).Value = "INSTITUTO DE TURISMO DEL CHACO"
    Cells(41, 6).Value = "42"
    Cells(41, 8).Value = Cells(37, 3).Value
    Cells(42, 7).Value = "MINISTERIO DE DESARROLLO URBANO Y ORD. TERRITORIAL"
    Cells(42, 6).Value = "45"
    Cells(42, 8).Value = Cells(39, 3).Value
    Cells(43, 7).Value = "INSTITUTO DEL DEFENSOR DEL PUEBLO"
    Cells(43, 6).Value = "47"
    Cells(43, 8).Value = Cells(41, 3).Value
    Cells(44, 7).Value = "INSTITUTO DEL DEPORTE"
    Cells(44, 6).Value = "48"
    Cells(44, 8).Value = Cells(42, 3).Value
    Cells(45, 7).Value = "MEC. PROV. DE PREV. DE LA TORTURA Y OTROS TRATOS"
    Cells(45, 6).Value = "50"
    Cells(45, 8).Value = Cells(44, 3).Value
    Cells(46, 7).Value = "INST. DESARROLLO RURAL Y AGRICULTURA  FLIAR"
    Cells(46, 6).Value = "52"
    Cells(46, 8).Value = Cells(46, 3).Value
    Cells(47, 7).Value = "SECRETARIA DE MUNICIPIOS Y CIUDADES"
    Cells(47, 6).Value = "53"
    Cells(47, 8).Value = Cells(47, 3).Value
    Cells(48, 7).Value = "PUERTO LAS PALMAS en creacion (D. 1094/18 - D. 2963/19, depende de la 23)"
    Cells(48, 6).Value = "55"
    Cells(48, 8).Value = Cells(49, 3).Value
    Cells(49, 6).Value = "TOTAL"
    
     Cells(49, 8).Value = 0
    
    For i = 23 To 48
     Cells(49, 8).Value = Cells(49, 8).Value + Cells(i, 8).Value
    Next i
    
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    
End Sub

Sub Total_por_Jur_No_Tratadas()

    Cells(53, 6).Value = "NO TRATADAS POR LOS DOS DEPARTAMENTOS"
    Cells(54, 6).Value = "JUR"
    Cells(54, 7).Value = "DENOMINACIÓN"
    Cells(54, 8).Value = "LIQUIDADO"
  
    Cells(55, 7).Value = "TRIBUNAL DE CUENTAS"
    Cells(55, 6).Value = "8"
    Cells(55, 8).Value = Cells(9, 3).Value
    Cells(56, 7).Value = "PODER JUDICIAL"
    Cells(56, 6).Value = "9"
    Cells(56, 8).Value = Cells(10, 3).Value
    Cells(57, 7).Value = "TRIBUNAL ELECTORAL"
    Cells(57, 6).Value = "17"
    Cells(57, 8).Value = Cells(17, 3).Value
    Cells(58, 7).Value = "FISCALIA DE ESTADO"
    Cells(58, 6).Value = "18"
    Cells(58, 8).Value = Cells(18, 3).Value
    Cells(59, 7).Value = "FISCALIA DE INVESTIGACIONES ADMINISTRATIVAS"
    Cells(59, 6).Value = "27"
    Cells(59, 8).Value = Cells(26, 3).Value
    Cells(60, 7).Value = "FONDO ESPECIAL DE RETIROS VOLUNTARIOS"
    Cells(60, 6).Value = "31"
    Cells(60, 8).Value = Cells(29, 3).Value
    Cells(61, 7).Value = "CONSEJO DE LA MAGISTRATURA Y JURADO ENJUICIAMIENTO"
    Cells(61, 6).Value = "40"
    Cells(61, 8).Value = Cells(36, 3).Value
    Cells(62, 7).Value = "FONDO ESPECIAL RETIRO  LEY 6636                   "
    Cells(62, 6).Value = "44"
    Cells(62, 8).Value = Cells(38, 3).Value
    Cells(63, 6).Value = "TOTAL"
    
    For i = 55 To 62
     Cells(63, 8).Value = Cells(63, 8).Value + Cells(i, 8).Value
    Next i
    
       
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"

End Sub

