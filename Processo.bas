Attribute VB_Name = "Processo"
Sub AtualizarTodasPlanilhas()
    Call MainEstatores
    Call MainRotorGaiola
    Call MainRotorBobinado
End Sub

Sub MainEstatores()
    Call DeclararVariaveis(1)
    Call ConectarSAP
    Call AnalisarLinhasMaterial
End Sub

Sub MainRotorGaiola()
    Call DeclararVariaveis(2)
    Call ConectarSAP
    Call AnalisarLinhasMaterial
End Sub

Sub MainRotorBobinado()
    Call DeclararVariaveis(3)
    Call ConectarSAP
    Call AnalisarLinhasMaterial
End Sub

Sub AnalisarLinhasMaterial()
    For Linha = 11 To 9999
        DataReal = PegarCelula(Linha, 3)
        Material = PegarCelula(Linha, 1)
        Ordem = PegarCelula(Linha, 2)
        
        If Ordem <> "" Then
            If Material = "" Then
                Call RemoverLinha(Linha)
                Call PegarDadosOrdem(Ordem)
            End If
        End If
        Call RestaurarVariaveis
    Next Linha
End Sub

Sub InserirPlanilha()
    Call InserirLinha
    Call DarValorCelula(11, 1, MaterialFaltante)
    Call DarValorCelula(11, 2, Ordem)
    Call DarValorCelula(11, 3, DataReal)
    Call DarValorCelula(11, 4, DataPlanejada)
    Call DarValorCelula(11, 5, SecaoCausadora)
    Call DarValorCelula(11, 6, Projeto)
    Call DarValorCelula(11, 7, DescricaoMaterial)
    Call DarValorCelula(11, 8, Fornecedor)
    Call DarValorCelula(11, 9, StatusComponente)
    Call PintarLinha
End Sub

Sub RestaurarVariaveis()
    Call RestaurarVariaveisSAP
    Call RestaurarVariaveisPlanilha
End Sub




