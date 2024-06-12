Attribute VB_Name = "SAP"
Sub PegarDadosOrdem(Ordem)
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nco03"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = Ordem
    session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").caretPosition = 8
    session.findById("wnd[0]/tbar[1]/btn[6]").press
    Call BuscarDadosTelaComponenteSAP
End Sub

Function BuscarMaterialFaltante(LinhaTabelaComponentes)
    BuscarMaterialFaltante = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-MATNR[1," & LinhaTabelaComponentes & "]").Text
End Function

Function BuscarDataPlanejada()
    session.findById("wnd[0]/tbar[1]/btn[5]").press
    DataPlanejadaTabela = session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-SSAVD[1,0]").Text
    session.findById("wnd[0]/tbar[1]/btn[6]").press
    BuscarDataPlanejada = DataPlanejadaTabela
End Function

Function BuscarSecaoCausadora(LinhaTabelaComponentes)
    session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-MATNR[1," & LinhaTabelaComponentes & "]").SetFocus
    session.findById("wnd[0]").sendVKey 2
    SecaoCausadoraTabela = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").Text
    session.findById("wnd[0]/tbar[0]/btn[15]").press
    BuscarSecaoCausadora = SecaoCausadoraTabela
End Function

Function BuscarProjeto()
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    session.findById("wnd[0]/usr/tabsTABSTRIP_0115/tabpKOAL").Select
    ProjetoTabela = session.findById("wnd[0]/usr/tabsTABSTRIP_0115/tabpKOAL/ssubSUBSCR_0115:SAPLCOKO1:0140/ctxtAFPOD-PROJN").Text
    session.findById("wnd[0]/tbar[1]/btn[6]").press
    BuscarProjeto = ProjetoTabela
End Function

Function BuscarDescricaoMaterial(LinhaTabelaComponentes)
    BuscarDescricaoMaterial = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MATXT[2," & LinhaTabelaComponentes & "]").Text
End Function

Function BuscarFornecedorPorMRP(MRP)
    Linhas = ContarLinhas(ListaFornecedores)
    For Linha = 2 To Linhas
    
        ValorMRP = ListaFornecedores.Cells(Linha, 1)
        
        If CDec(MRP) = CDec(ValorMRP) Then
            BuscarFornecedorPorMRP = ListaFornecedores.Cells(Linha, 2)
            Exit Function
        End If
        
    Next Linha
End Function

Function BuscarDadosTelaComponenteSAP()
    For LinhaTabelaComponentes = 0 To 29
    
        OperacaoNumero = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-VORNR[6," & LinhaTabelaComponentes & "]").Text
        If Not IsNumeric(OperacaoNumero) Then
            Exit For
        End If
        
        OperacaoNumero = CDec(OperacaoNumero)
        If OperacaoNumero <= numero_ate_onde_operacao_vai Then
            NumeroAtrasado = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-DVMENG[11," & LinhaTabelaComponentes & "]").Text
            QuantidadeRetirada = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-DENMNG[12," & LinhaTabelaComponentes & "]").Text
            
            MaterialFaltante = BuscarMaterialFaltante(LinhaTabelaComponentes)
            DataPlanejada = BuscarDataPlanejada()
            Projeto = BuscarProjeto()
            DescricaoMaterial = BuscarDescricaoMaterial(LinhaTabelaComponentes)
        
            If QuantidadeRetirada > NumeroAtrasado Then
                StatusComponente = "MATERIAL UTILIZADO"
                Fornecedor = ""
            End If
            
            If NumeroAtrasado > 0 Then
                StatusComponente = "Est� no estoque"
                Fornecedor = ""
            End If
                        
            If NumeroAtrasado = 0 Then
                If QuantidadeRetirada = 0 Then
                    SecaoCausadora = BuscarSecaoCausadora(LinhaTabelaComponentes)
                    Fornecedor = BuscarFornecedorPorMRP(SecaoCausadora)
                    StatusComponente = "Faltando no estoque"
                End If
            End If
            
            Call InserirPlanilha
            
        End If
    Next LinhaTabelaComponentes
End Function
