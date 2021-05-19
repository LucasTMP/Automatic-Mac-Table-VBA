Sub Limpar_DHCP()

Sheets("Controle Rede Local").Range("tabela_dhcp").Select
Selection.ClearContents
Selection.ClearFormats
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
End Sub
