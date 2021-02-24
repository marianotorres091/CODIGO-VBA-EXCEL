Attribute VB_Name = "Módulo1"
Sub Total_por_Jur_x_mes_cpto_233()
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
    
    Cells(2, 3).Value = "MAR 2015"
    Cells(2, 4).Value = "ABR 2015"
    Cells(2, 5).Value = "MAY 2015"
    Cells(2, 6).Value = "JUN 2015 + SAC"
    Cells(2, 7).Value = "JUL 2015"
    Cells(2, 8).Value = "AGOS 2015"
    Cells(2, 9).Value = "SEP 2015"
    Cells(2, 10).Value = "OCT 2015"
    Cells(2, 11).Value = "NOV 2015"
    Cells(2, 12).Value = "DIC 2015C + SAC"
    Cells(2, 13).Value = "ENE 2016"
    Cells(2, 14).Value = "FEB 2016"
    Cells(2, 15).Value = "MAR 2016"
    Cells(2, 16).Value = "ABR 2016"
    Cells(2, 17).Value = "MAY 2016"
    Cells(2, 18).Value = "JUN 2016 + SAC"
    Cells(2, 19).Value = "JUL 2016"
    Cells(2, 20).Value = "AGOS 2016"
    Cells(2, 21).Value = "SEP 2016"
    Cells(2, 22).Value = "OCT 2016"
    Cells(2, 23).Value = "DIC 2016 + SAC"
    
    
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
      Sheets("REPORTE").Cells(i, 4).Value = 0
      Sheets("REPORTE").Cells(i, 5).Value = 0
      Sheets("REPORTE").Cells(i, 6).Value = 0
      Sheets("REPORTE").Cells(i, 7).Value = 0
      Sheets("REPORTE").Cells(i, 8).Value = 0
      Sheets("REPORTE").Cells(i, 9).Value = 0
      Sheets("REPORTE").Cells(i, 10).Value = 0
      Sheets("REPORTE").Cells(i, 11).Value = 0
      Sheets("REPORTE").Cells(i, 12).Value = 0
      Sheets("REPORTE").Cells(i, 13).Value = 0
      Sheets("REPORTE").Cells(i, 14).Value = 0
      Sheets("REPORTE").Cells(i, 15).Value = 0
      Sheets("REPORTE").Cells(i, 16).Value = 0
      Sheets("REPORTE").Cells(i, 17).Value = 0
      Sheets("REPORTE").Cells(i, 18).Value = 0
      Sheets("REPORTE").Cells(i, 19).Value = 0
      Sheets("REPORTE").Cells(i, 20).Value = 0
      Sheets("REPORTE").Cells(i, 21).Value = 0
      Sheets("REPORTE").Cells(i, 22).Value = 0
      Sheets("REPORTE").Cells(i, 23).Value = 0
      
      
       For j = 2 To nFilas
       
          If Sheets("PERIODO 32015 AL 122016").Cells(j, 1).Value = Sheets("REPORTE").Cells(i, 1).Value Then
            
            mes = Sheets("PERIODO 32015 AL 122016").Cells(j, 16).Value
            Select Case mes
              Case Is = 32015
               Sheets("REPORTE").Cells(i, 3).Value = Sheets("REPORTE").Cells(i, 3).Value + Sheets("PERIODO 32015 AL 122016").Cells(j, 15).Value
              Case Is = 42015
               Sheets("REPORTE").Cells(i, 4).Value = Sheets("REPORTE").Cells(i, 4).Value + Sheets("PERIODO 32015 AL 122016").Cells(j, 15).Value
              Case Is = 52015
               Sheets("REPORTE").Cells(i, 5).Value = Sheets("REPORTE").Cells(i, 5).Value + Sheets("PERIODO 32015 AL 122016").Cells(j, 15).Value
              Case Is = 62015
               Sheets("REPORTE").Cells(i, 6).Value = Sheets("REPORTE").Cells(i, 6).Value + Sheets("PERIODO 32015 AL 122016").Cells(j, 15).Value + Sheets("PERIODO 32015 AL 122016").Cells(j, 17).Value
              Case Is = 72015
               Sheets("REPORTE").Cells(i, 7).Value = Sheets("REPORTE").Cells(i, 7).Value + Sheets("PERIODO 32015 AL 122016").Cells(j, 15).Value
              Case Is = 82015
               Sheets("REPORTE").Cells(i, 8).Value = Sheets("REPORTE").Cells(i, 8).Value + Sheets("PERIODO 32015 AL 122016").Cells(j, 15).Value
              Case Is = 92015
               Sheets("REPORTE").Cells(i, 9).Value = Sheets("REPORTE").Cells(i, 9).Value + Sheets("PERIODO 32015 AL 122016").Cells(j, 15).Value
              Case Is = 102015
               Sheets("REPORTE").Cells(i, 10).Value = Sheets("REPORTE").Cells(i, 10).Value + Sheets("PERIODO 32015 AL 122016").Cells(j, 15).Value
              Case Is = 112015
               Sheets("REPORTE").Cells(i, 11).Value = Sheets("REPORTE").Cells(i, 11).Value + Sheets("PERIODO 32015 AL 122016").Cells(j, 15).Value
              Case Is = 122015
               Sheets("REPORTE").Cells(i, 12).Value = Sheets("REPORTE").Cells(i, 12).Value + Sheets("PERIODO 32015 AL 122016").Cells(j, 15).Value + Sheets("PERIODO 32015 AL 122016").Cells(j, 17).Value
              Case Is = 12016
               Sheets("REPORTE").Cells(i, 13).Value = Sheets("REPORTE").Cells(i, 13).Value + Sheets("PERIODO 32015 AL 122016").Cells(j, 15).Value
              Case Is = 22016
               Sheets("REPORTE").Cells(i, 14).Value = Sheets("REPORTE").Cells(i, 14).Value + Sheets("PERIODO 32015 AL 122016").Cells(j, 15).Value
              Case Is = 32016
               Sheets("REPORTE").Cells(i, 15).Value = Sheets("REPORTE").Cells(i, 15).Value + Sheets("PERIODO 32015 AL 122016").Cells(j, 15).Value
              Case Is = 42016
               Sheets("REPORTE").Cells(i, 16).Value = Sheets("REPORTE").Cells(i, 16).Value + Sheets("PERIODO 32015 AL 122016").Cells(j, 15).Value
              Case Is = 52016
               Sheets("REPORTE").Cells(i, 17).Value = Sheets("REPORTE").Cells(i, 17).Value + Sheets("PERIODO 32015 AL 122016").Cells(j, 15).Value
              Case Is = 62016
               Sheets("REPORTE").Cells(i, 18).Value = Sheets("REPORTE").Cells(i, 18).Value + Sheets("PERIODO 32015 AL 122016").Cells(j, 15).Value + Sheets("PERIODO 32015 AL 122016").Cells(j, 17).Value
              Case Is = 72016
               Sheets("REPORTE").Cells(i, 19).Value = Sheets("REPORTE").Cells(i, 19).Value + Sheets("PERIODO 32015 AL 122016").Cells(j, 15).Value
              Case Is = 82016
               Sheets("REPORTE").Cells(i, 20).Value = Sheets("REPORTE").Cells(i, 20).Value + Sheets("PERIODO 32015 AL 122016").Cells(j, 15).Value
              Case Is = 92016
               Sheets("REPORTE").Cells(i, 21).Value = Sheets("REPORTE").Cells(i, 21).Value + Sheets("PERIODO 32015 AL 122016").Cells(j, 15).Value
              Case Is = 102016
               Sheets("REPORTE").Cells(i, 22).Value = Sheets("REPORTE").Cells(i, 22).Value + Sheets("PERIODO 32015 AL 122016").Cells(j, 15).Value
              Case Is = 122016
               Sheets("REPORTE").Cells(i, 23).Value = Sheets("REPORTE").Cells(i, 23).Value + Sheets("PERIODO 32015 AL 122016").Cells(j, 15).Value + Sheets("PERIODO 32015 AL 122016").Cells(j, 17).Value
            End Select
          End If
          
        Next j
        
       Sheets("REPORTE").Cells(50, 3).Value = Sheets("REPORTE").Cells(50, 3).Value + Sheets("REPORTE").Cells(i, 3).Value
       Sheets("REPORTE").Cells(50, 4).Value = Sheets("REPORTE").Cells(50, 4).Value + Sheets("REPORTE").Cells(i, 4).Value
       Sheets("REPORTE").Cells(50, 5).Value = Sheets("REPORTE").Cells(50, 5).Value + Sheets("REPORTE").Cells(i, 5).Value
       Sheets("REPORTE").Cells(50, 6).Value = Sheets("REPORTE").Cells(50, 6).Value + Sheets("REPORTE").Cells(i, 6).Value
       Sheets("REPORTE").Cells(50, 7).Value = Sheets("REPORTE").Cells(50, 7).Value + Sheets("REPORTE").Cells(i, 7).Value
       Sheets("REPORTE").Cells(50, 8).Value = Sheets("REPORTE").Cells(50, 8).Value + Sheets("REPORTE").Cells(i, 8).Value
       Sheets("REPORTE").Cells(50, 9).Value = Sheets("REPORTE").Cells(50, 9).Value + Sheets("REPORTE").Cells(i, 9).Value
       Sheets("REPORTE").Cells(50, 10).Value = Sheets("REPORTE").Cells(50, 10).Value + Sheets("REPORTE").Cells(i, 10).Value
       Sheets("REPORTE").Cells(50, 11).Value = Sheets("REPORTE").Cells(50, 11).Value + Sheets("REPORTE").Cells(i, 11).Value
       Sheets("REPORTE").Cells(50, 12).Value = Sheets("REPORTE").Cells(50, 12).Value + Sheets("REPORTE").Cells(i, 12).Value
       Sheets("REPORTE").Cells(50, 13).Value = Sheets("REPORTE").Cells(50, 13).Value + Sheets("REPORTE").Cells(i, 13).Value
       Sheets("REPORTE").Cells(50, 14).Value = Sheets("REPORTE").Cells(50, 14).Value + Sheets("REPORTE").Cells(i, 14).Value
       Sheets("REPORTE").Cells(50, 15).Value = Sheets("REPORTE").Cells(50, 15).Value + Sheets("REPORTE").Cells(i, 15).Value
       Sheets("REPORTE").Cells(50, 16).Value = Sheets("REPORTE").Cells(50, 16).Value + Sheets("REPORTE").Cells(i, 16).Value
       Sheets("REPORTE").Cells(50, 17).Value = Sheets("REPORTE").Cells(50, 17).Value + Sheets("REPORTE").Cells(i, 17).Value
       Sheets("REPORTE").Cells(50, 18).Value = Sheets("REPORTE").Cells(50, 18).Value + Sheets("REPORTE").Cells(i, 18).Value
       Sheets("REPORTE").Cells(50, 19).Value = Sheets("REPORTE").Cells(50, 19).Value + Sheets("REPORTE").Cells(i, 19).Value
       Sheets("REPORTE").Cells(50, 20).Value = Sheets("REPORTE").Cells(50, 20).Value + Sheets("REPORTE").Cells(i, 20).Value
       Sheets("REPORTE").Cells(50, 21).Value = Sheets("REPORTE").Cells(50, 21).Value + Sheets("REPORTE").Cells(i, 21).Value
       Sheets("REPORTE").Cells(50, 22).Value = Sheets("REPORTE").Cells(50, 22).Value + Sheets("REPORTE").Cells(i, 22).Value
       Sheets("REPORTE").Cells(50, 23).Value = Sheets("REPORTE").Cells(50, 23).Value + Sheets("REPORTE").Cells(i, 23).Value
       
       Sheets("REPORTE").Cells(i, 25).Value = Sheets("REPORTE").Cells(i, 25).Value + Sheets("REPORTE").Cells(i, 3).Value
       Sheets("REPORTE").Cells(i, 25).Value = Sheets("REPORTE").Cells(i, 25).Value + Sheets("REPORTE").Cells(i, 4).Value
       Sheets("REPORTE").Cells(i, 25).Value = Sheets("REPORTE").Cells(i, 25).Value + Sheets("REPORTE").Cells(i, 5).Value
       Sheets("REPORTE").Cells(i, 25).Value = Sheets("REPORTE").Cells(i, 25).Value + Sheets("REPORTE").Cells(i, 6).Value
       Sheets("REPORTE").Cells(i, 25).Value = Sheets("REPORTE").Cells(i, 25).Value + Sheets("REPORTE").Cells(i, 7).Value
       Sheets("REPORTE").Cells(i, 25).Value = Sheets("REPORTE").Cells(i, 25).Value + Sheets("REPORTE").Cells(i, 8).Value
       Sheets("REPORTE").Cells(i, 25).Value = Sheets("REPORTE").Cells(i, 25).Value + Sheets("REPORTE").Cells(i, 9).Value
       Sheets("REPORTE").Cells(i, 25).Value = Sheets("REPORTE").Cells(i, 25).Value + Sheets("REPORTE").Cells(i, 10).Value
       Sheets("REPORTE").Cells(i, 25).Value = Sheets("REPORTE").Cells(i, 25).Value + Sheets("REPORTE").Cells(i, 11).Value
       Sheets("REPORTE").Cells(i, 25).Value = Sheets("REPORTE").Cells(i, 25).Value + Sheets("REPORTE").Cells(i, 12).Value
       Sheets("REPORTE").Cells(i, 25).Value = Sheets("REPORTE").Cells(i, 25).Value + Sheets("REPORTE").Cells(i, 13).Value
       Sheets("REPORTE").Cells(i, 25).Value = Sheets("REPORTE").Cells(i, 25).Value + Sheets("REPORTE").Cells(i, 14).Value
       Sheets("REPORTE").Cells(i, 25).Value = Sheets("REPORTE").Cells(i, 25).Value + Sheets("REPORTE").Cells(i, 15).Value
       Sheets("REPORTE").Cells(i, 25).Value = Sheets("REPORTE").Cells(i, 25).Value + Sheets("REPORTE").Cells(i, 16).Value
       Sheets("REPORTE").Cells(i, 25).Value = Sheets("REPORTE").Cells(i, 25).Value + Sheets("REPORTE").Cells(i, 17).Value
       Sheets("REPORTE").Cells(i, 25).Value = Sheets("REPORTE").Cells(i, 25).Value + Sheets("REPORTE").Cells(i, 18).Value
       Sheets("REPORTE").Cells(i, 25).Value = Sheets("REPORTE").Cells(i, 25).Value + Sheets("REPORTE").Cells(i, 19).Value
       Sheets("REPORTE").Cells(i, 25).Value = Sheets("REPORTE").Cells(i, 25).Value + Sheets("REPORTE").Cells(i, 20).Value
       Sheets("REPORTE").Cells(i, 25).Value = Sheets("REPORTE").Cells(i, 25).Value + Sheets("REPORTE").Cells(i, 21).Value
       Sheets("REPORTE").Cells(i, 25).Value = Sheets("REPORTE").Cells(i, 25).Value + Sheets("REPORTE").Cells(i, 22).Value
       Sheets("REPORTE").Cells(i, 25).Value = Sheets("REPORTE").Cells(i, 25).Value + Sheets("REPORTE").Cells(i, 23).Value
        
    Next i
       
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
    
    Application.StatusBar = False
    
    Exit Sub

End Sub


