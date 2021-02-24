Attribute VB_Name = "Módulo1"
Sub Totales_Cpto_Jur()
    Dim nFilas As Long
    Dim filaTotal As Long
    Dim rango As Range
    Dim wsTotal As Excel.Worksheet
    Dim i As Long
    Dim ultJur As Integer
    Dim ultCpto As Integer
    Dim total As Double
    
    
    Application.DisplayAlerts = False
    Worksheets.Add
    ActiveSheet.Name = "Total Cpto"
    Application.DisplayAlerts = True
    
    Set wsTotal = Worksheets("Total Cpto")
    
    'Regresa el control a la hoja de origen
    Sheets("Hoja1").Select
    
    'Calcular el número de filas de la hoja actual
    Set rango = ActiveSheet.UsedRange
    nFilas = rango.Rows.Count
    
    MsgBox "Debe estar ordenado por JUR + CPTO.", , "Atención!!"
    MsgBox "NO ACUMULA LOS IMPORTES DE LOS CPTOS MAYORES A 400 !!!"
    
    'Encabezado Hoja Totales
    wsTotal.Cells(1, 1).Value = "JUR"
    wsTotal.Cells(1, 2).Value = "CPTO"
    wsTotal.Cells(1, 3).Value = "Descripción"
    wsTotal.Cells(1, 4).Value = "Importe"
    filaTotal = 2
    
    ultJur = Cells(2, 3).Value
    ultCpto = Cells(2, 9).Value
    descripcion = Cells(2, 9).Value
    
     Select Case descripcion
     Case Is = 1
      descripcion = "SUELDO BASICO"
     Case Is = 7
      descripcion = "COMPLEMENTO BASICO"
     Case Is = 17
      descripcion = "TIEMPO COMPLETO EN BASICO"
     Case Is = 25
      descripcion = "PRESTACIONES PROF. COMPL."
     Case Is = 29
      descripcion = "DEDICACION LEY 3490"
     Case Is = 30
      descripcion = "RES: FUNCIONAL LEY 3490"
     Case Is = 44
      descripcion = "BONIFICACION ESPECIAL LEY 6865"
     Case Is = 45
      descripcion = "BONIFICACION RESPONSABILIDAD FUNCIONAL AUT. SUPERIOR"
     Case Is = 47
      descripcion = "BONIF.RESPONSABILIDA"
     Case Is = 120
      descripcion = "DEDICACION"
     Case Is = 121
      descripcion = "DEDICACION AUTORIDADES SUPERIORES"
     Case Is = 122
      descripcion = "BONIFICACION DEDICACION LEY 6655"
     Case Is = 126
      descripcion = "DEDICACION EXCLUSIVA"
     Case Is = 128
      descripcion = "BONIFICACION POR DEDICACION"
     Case Is = 129
      descripcion = "SERVICIOS ESPECIALES"
     Case Is = 130
      descripcion = "COMPENSACION PROLONG.HORAR."
     Case Is = 134
      descripcion = "DEDICACION EXCLUSIVA LEY 5315"
     Case Is = 138
      descripcion = "ADICIONAL PARTICULAR POR MAYOR DEDICACION"
     Case Is = 142
      descripcion = "BONIFICACION ESPECIAL INSP. GRAL PERS JUR."
     Case Is = 150
      descripcion = "SEGURIDAD - BONIFICACION ESPECIAL"
     Case Is = 170
      descripcion = "FUNC.DIFERENCIADA Y/O ESP."
     Case Is = 190
      descripcion = "ATENCION PRACTICANTES"
     Case Is = 198
      descripcion = "TIEMPO MINIMO CUMPLIDO"
     Case Is = 203
      descripcion = "HORAS EXTRAORDINARIAS LOTERIA CHAQUEÑA"
     Case Is = 204
      descripcion = "PRESENTISMO"
     Case Is = 210
      descripcion = "INCOMPATIBILIDAD"
     Case Is = 212
      descripcion = "TITULO SECUNDARIO"
     Case Is = 213
      descripcion = "TITULO UNIVERSITARIO"
     Case Is = 215
      descripcion = "TITULO TERCIARIO"
     Case Is = 216
      descripcion = "ADICIONAL COMPLEMENTARIO"
     Case Is = 217
      descripcion = "HEROES DE MALVINAS"
     Case Is = 219
      descripcion = "ZONA DESFAVORABLE DESARROLLO SOCIAL"
     Case Is = 220
      descripcion = "ZONA DESFAVORABLE"
     Case Is = 221
      descripcion = "ZONA UNIDADES ESPECIALES"
     Case Is = 222
      descripcion = "DESARRAIGO"
     Case Is = 223
      descripcion = "RIESGO"
     Case Is = 225
      descripcion = "ADICIONAL POR TAREA ESPECIFICA"
     Case Is = 230
      descripcion = "PERM.EN EL CARGO"
     Case Is = 233
      descripcion = "BONIF. LEY 6655"
     Case Is = 234
      descripcion = "DECRETO 483"
     Case Is = 239
      descripcion = "BONIFICACION SERVICIOS ADMINISTRATIVOS ADICIONALES"
     Case Is = 242
      descripcion = "RIESGO P/EXPLOSIVOS"
     Case Is = 244
      descripcion = "RIESGO PROFESIONAL"
     Case Is = 246
      descripcion = "RIESGO DE SALUD"
     Case Is = 248
      descripcion = "INSALUBRIDAD"
     Case Is = 250
      descripcion = "RIESGO DE VIDA"
     Case Is = 251
      descripcion = "BONIFICACION P ASISTENCIA EXCLUSIVA AL LEGISLADOR"
     Case Is = 252
      descripcion = "BONIFICACION RIESGO DE VIDA ESPECIAL"
     Case Is = 253
      descripcion = "BONIF. TITULO PRIMARIO Y BASICO"
     Case Is = 254
      descripcion = "RIESGO VISUAL"
     Case Is = 255
      descripcion = "BONIF. POR PERMANENCIA EN LA ESTRUCTURA"
     Case Is = 265
      descripcion = "ADICIONAL REMUNERATIVO PUERTO"
     Case Is = 267
      descripcion = "FONDO ESTIMULO PRODUCTIVO"
     Case Is = 270
      descripcion = "CAPACITACION POLICIAL"
     Case Is = 273
      descripcion = "BONIFICACION P AGENTES DE ESTAB. SANITARIOS"
     Case Is = 274
      descripcion = "BONIFICACION AUXILIARES DE ENFERMERIA"
     Case Is = 275
      descripcion = "GUARDIA ACTIVA"
     Case Is = 277
      descripcion = "GUARDIA PASIVA"
     Case Is = 278
      descripcion = "RECONOCIMIENTOS MEDICOS"
     Case Is = 279
      descripcion = "JUNTAS MEDICAS"
     Case Is = 280
      descripcion = "INCOMPATIBILIDAD PARCIAL"
     Case Is = 281
      descripcion = "INCOMPATIBILIDAD DCCION. AERONAUTICA"
     Case Is = 282
      descripcion = "JORNADA NOCTURNA"
     Case Is = 299
      descripcion = "BONIFICACION ESPECIAL"
     Case Is = 302
      descripcion = "SUMA FIJA"
     Case Is = 306
      descripcion = "SUPLEMENTO POR FUNCION"
     Case Is = 307
      descripcion = "BONIFICACION ESPECIAL SEC.INVERSIONES"
     Case Is = 308
      descripcion = "RIESGO DE CAJA"
     Case Is = 309
      descripcion = "FONDO FORTALECIMIENTO DE GESTION INSTITUCIONAL"
     Case Is = 312
      descripcion = "BONIFICACION ESPECIAL COLONIZACION"
     Case Is = 333
      descripcion = "SUPLEMENTO POR FUNCION"
     Case Else
      descripcion = "FALTA DESCRIPCION"
     End Select
     
    importe = 0
    total = 0
    
    limite = nFilas
    
    For i = 2 To limite
     Application.StatusBar = Format(i / limite, "0.0%") & "Completo"
        If Cells(i, 9).Value < 400 Then
            If ultCpto = Cells(i, 9).Value And ultJur = Cells(i, 3).Value Then
                If Cells(i, 10).Value = 2 Then
                    importe = importe - Cells(i, 12).Value
                Else
                    importe = importe + Cells(i, 12).Value
                End If
            Else
                wsTotal.Cells(filaTotal, 1).Value = ultJur
                wsTotal.Cells(filaTotal, 2).Value = ultCpto
                wsTotal.Cells(filaTotal, 3).Value = descripcion
                wsTotal.Cells(filaTotal, 4).Value = importe
                total = total + importe
                filaTotal = filaTotal + 1
                If ultJur <> Cells(i, 3).Value Then
                    wsTotal.Cells(filaTotal, 3).Value = "TOTAL" & " " & "JUR" & " " & ultJur
                    wsTotal.Cells(filaTotal, 4).Value = total
                    total = 0
                    filaTotal = filaTotal + 1
                End If
                ultCpto = Cells(i, 9).Value
                ultJur = Cells(i, 3).Value
                descripcion = Cells(i, 9).Value
                Select Case descripcion
                    Case Is = 1
                     descripcion = "SUELDO BASICO"
                    Case Is = 7
                     descripcion = "COMPLEMENTO BASICO"
                    Case Is = 17
                     descripcion = "TIEMPO COMPLETO EN BASICO"
                    Case Is = 25
                     descripcion = "PRESTACIONES PROF. COMPL."
                    Case Is = 29
                     descripcion = "DEDICACION LEY 3490"
                    Case Is = 30
                     descripcion = "RES: FUNCIONAL LEY 3490"
                    Case Is = 44
                     descripcion = "BONIFICACION ESPECIAL LEY 6865"
                    Case Is = 45
                     descripcion = "BONIFICACION RESPONSABILIDAD FUNCIONAL AUT. SUPERIOR"
                    Case Is = 47
                     descripcion = "BONIF.RESPONSABILIDA"
                    Case Is = 120
                     descripcion = "DEDICACION"
                    Case Is = 121
                     descripcion = "DEDICACION AUTORIDADES SUPERIORES"
                    Case Is = 122
                     descripcion = "BONIFICACION DEDICACION LEY 6655"
                    Case Is = 126
                     descripcion = "DEDICACION EXCLUSIVA"
                    Case Is = 128
                     descripcion = "BONIFICACION POR DEDICACION"
                    Case Is = 129
                     descripcion = "SERVICIOS ESPECIALES"
                    Case Is = 130
                     descripcion = "COMPENSACION PROLONG.HORAR."
                    Case Is = 134
                     descripcion = "DEDICACION EXCLUSIVA LEY 5315"
                    Case Is = 138
                     descripcion = "ADICIONAL PARTICULAR POR MAYOR DEDICACION"
                    Case Is = 142
                     descripcion = "BONIFICACION ESPECIAL INSP. GRAL PERS JUR."
                    Case Is = 150
                     descripcion = "SEGURIDAD - BONIFICACION ESPECIAL"
                    Case Is = 170
                     descripcion = "FUNC.DIFERENCIADA Y/O ESP."
                    Case Is = 190
                     descripcion = "ATENCION PRACTICANTES"
                    Case Is = 198
                     descripcion = "TIEMPO MINIMO CUMPLIDO"
                    Case Is = 203
                     descripcion = "HORAS EXTRAORDINARIAS LOTERIA CHAQUEÑA"
                    Case Is = 204
                     descripcion = "PRESENTISMO"
                    Case Is = 210
                     descripcion = "INCOMPATIBILIDAD"
                    Case Is = 212
                     descripcion = "TITULO SECUNDARIO"
                    Case Is = 213
                     descripcion = "TITULO UNIVERSITARIO"
                    Case Is = 215
                     descripcion = "TITULO TERCIARIO"
                    Case Is = 216
                     descripcion = "ADICIONAL COMPLEMENTARIO"
                    Case Is = 217
                     descripcion = "HEROES DE MALVINAS"
                    Case Is = 219
                     descripcion = "ZONA DESFAVORABLE DESARROLLO SOCIAL"
                    Case Is = 220
                     descripcion = "ZONA DESFAVORABLE"
                    Case Is = 221
                     descripcion = "ZONA UNIDADES ESPECIALES"
                    Case Is = 222
                     descripcion = "DESARRAIGO"
                    Case Is = 223
                     descripcion = "RIESGO"
                    Case Is = 225
                     descripcion = "ADICIONAL POR TAREA ESPECIFICA"
                    Case Is = 230
                     descripcion = "PERM.EN EL CARGO"
                    Case Is = 233
                     descripcion = "BONIF. LEY 6655"
                    Case Is = 234
                     descripcion = "DECRETO 483"
                    Case Is = 239
                     descripcion = "BONIFICACION SERVICIOS ADMINISTRATIVOS ADICIONALES"
                    Case Is = 242
                     descripcion = "RIESGO P/EXPLOSIVOS"
                    Case Is = 244
                     descripcion = "RIESGO PROFESIONAL"
                    Case Is = 246
                     descripcion = "RIESGO DE SALUD"
                    Case Is = 248
                     descripcion = "INSALUBRIDAD"
                    Case Is = 250
                     descripcion = "RIESGO DE VIDA"
                    Case Is = 251
                     descripcion = "BONIFICACION P ASISTENCIA EXCLUSIVA AL LEGISLADOR"
                    Case Is = 252
                     descripcion = "BONIFICACION RIESGO DE VIDA ESPECIAL"
                    Case Is = 253
                     descripcion = "BONIF. TITULO PRIMARIO Y BASICO"
                    Case Is = 254
                     descripcion = "RIESGO VISUAL"
                    Case Is = 255
                     descripcion = "BONIF. POR PERMANENCIA EN LA ESTRUCTURA"
                    Case Is = 265
                     descripcion = "ADICIONAL REMUNERATIVO PUERTO"
                    Case Is = 267
                     descripcion = "FONDO ESTIMULO PRODUCTIVO"
                    Case Is = 270
                     descripcion = "CAPACITACION POLICIAL"
                    Case Is = 273
                     descripcion = "BONIFICACION P AGENTES DE ESTAB. SANITARIOS"
                    Case Is = 274
                     descripcion = "BONIFICACION AUXILIARES DE ENFERMERIA"
                    Case Is = 275
                     descripcion = "GUARDIA ACTIVA"
                    Case Is = 277
                     descripcion = "GUARDIA PASIVA"
                    Case Is = 278
                     descripcion = "RECONOCIMIENTOS MEDICOS"
                    Case Is = 279
                     descripcion = "JUNTAS MEDICAS"
                    Case Is = 280
                     descripcion = "INCOMPATIBILIDAD PARCIAL"
                    Case Is = 281
                     descripcion = "INCOMPATIBILIDAD DCCION. AERONAUTICA"
                    Case Is = 282
                     descripcion = "JORNADA NOCTURNA"
                    Case Is = 299
                     descripcion = "BONIFICACION ESPECIAL"
                    Case Is = 302
                     descripcion = "SUMA FIJA"
                    Case Is = 306
                     descripcion = "SUPLEMENTO POR FUNCION"
                    Case Is = 307
                     descripcion = "BONIFICACION ESPECIAL SEC.INVERSIONES"
                    Case Is = 308
                     descripcion = "RIESGO DE CAJA"
                    Case Is = 309
                     descripcion = "FONDO FORTALECIMIENTO DE GESTION INSTITUCIONAL"
                    Case Is = 312
                     descripcion = "BONIFICACION ESPECIAL COLONIZACION"
                    Case Is = 333
                     descripcion = "SUPLEMENTO POR FUNCION"
                    Case Else
                     descripcion = "FALTA DESCRIPCION"
                End Select
                importe = 0
                i = i - 1
            End If
        End If
    Next i
    wsTotal.Cells(filaTotal, 1).Value = ultJur
    wsTotal.Cells(filaTotal, 2).Value = ultCpto
    wsTotal.Cells(filaTotal, 3).Value = descripcion
    wsTotal.Cells(filaTotal, 4).Value = importe
    wsTotal.Cells(filaTotal + 1, 4).Value = total + importe
    wsTotal.Cells(filaTotal + 1, 3).Value = "TOTAL" & " " & "JUR" & " " & ultJur
    
    Application.StatusBar = False
    MsgBox "Se ha realizado con éxito la operación.", , "Finalizado"
End Sub


   
