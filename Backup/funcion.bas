Attribute VB_Name = "Módulo2"
Function IMPORTE_GUARDIA(couc, prof, horas)
    
    importe As Double
    
    If couc = "276" Then
        If prof = "A" Then
            importe = horas * 150
        Else
            If prof = "B" Then
                importe = horas * 140
            Else
                importe = horas * 85
            End If
        End If
    Else
        If couc = "275" Then
            If prof = "A" Then
                importe = horas * 100
            Else
                If prof = "B" Then
                    importe = horas * 90
                Else
                    importe = horas * 70
                End If
            End If
        Else
            importe = horas * 40
        End If
    End If
    
    IMPORTE_GUARDIA = importe
    
End Function
