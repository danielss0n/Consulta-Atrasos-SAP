Attribute VB_Name = "Planilha"
Function PegarCelula(Linha, Coluna)
    PegarCelula = ThisWorksheet.Cells(Linha, Coluna).Value
End Function

Function DarValorCelula(Linha, Coluna, Valor)
    ThisWorksheet.Cells(Linha, Coluna).Value = Valor
End Function

Function InserirLinha()
    ThisWorksheet.Rows(11).Insert
    ThisWorksheet.Rows(11).ClearFormats
    ThisWorksheet.Rows(11).Columns("A:I").Borders.LineStyle = xlContinuous
End Function

Function ContarLinhas(ws As Worksheet)
    ContarLinhas = ThisWorksheet.UsedRange.Rows.Count
End Function

Function RemoverLinha(Linha)
    ThisWorksheet.Rows(Linha).EntireRow.Delete
End Function

Function RestaurarVariaveisSAP()
    MaterialFaltante = ""
    DataPlanejada = ""
    SecaoCausadora = ""
    Projeto = ""
    DescricaoMaterial = ""
End Function

Function RestaurarVariaveisPlanilha()
    DataReal = ""
    Material = ""
    Ordem = ""
End Function

Function PintarLinha()
    With ThisWorksheet.Range("A11:I11").Interior
        Select Case StatusComponente
            Case "Faltando no estoque"
                .Color = RGB(255, 255, 150)
            Case "Est� no estoque"
                .Color = RGB(150, 255, 150)
            Case "MATERIAL UTILIZADO"
                .Color = RGB(150, 150, 255)
        End Select
    End With
End Function

