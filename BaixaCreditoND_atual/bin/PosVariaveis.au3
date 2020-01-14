;Padrao Nomeacao das Variaveis do Projeto
;Variaveis de Posicionamento da Tela
;Sx - Numero do Sistea + Numero da Tela + Identificao do Texto + x ou y
;1 - Atlys
;2 - Ngin
;3 - VivoNet
;Exemplo S1T1PesquisarX, S1T1PesquisarY


Global $sPesquisaPagamento[14]
Global $sReversaoPagamento[14]
Global $sPrintReversao[4]

Global Enum $eOpcaoPesquisaPgtoX, _
			$eOpcaoPesquisaPgtoY, _
			$eMaximizarTelaPgtoX, _
			$eMaximizarTelaPgtoY, _
			$eColorGridTLoteX	, _
			$eColorGridTLoteY	, _
			$eClickPaginarTLoteX, _
			$eClickPaginarTLoteY, _
			$eMapaPrintTLoteInicioX , _
			$eMapaPrintTLoteInicioY , _
			$eMapaPrintTLoteFimX , _
			$eMapaPrintTLoteFimY , _
			$eFecharTelaPgtoX, _
			$eFecharTelaPgtoY

 Global Enum $eRegContasX, _
			 $eRegContasY, _
			 $eRegContasFimX, _
			 $eRegContasFimY

;------------------------------------------------------
$sPrintReversao[$eRegContasX]			=	40
$sPrintReversao[$eRegContasY]			=	340
$sPrintReversao[$eRegContasFimX]		=	1190
$sPrintReversao[$eRegContasFimY]		=	476
;------------------------------------------------------


;------------------------------------------------------
$sPesquisaPagamento[$eOpcaoPesquisaPgtoX]	=	168
$sPesquisaPagamento[$eOpcaoPesquisaPgtoY]	=	65
;------------------------------------------------------

;------------------------------------------------------
$sPesquisaPagamento[$eMaximizarTelaPgtoX]	=	605
$sPesquisaPagamento[$eMaximizarTelaPgtoY]	=	100
;------------------------------------------------------

;------------------------------------------------------
$sPesquisaPagamento[$eColorGridTLoteX]	=	1316
$sPesquisaPagamento[$eColorGridTLoteY]	=	579
;------------------------------------------------------

;------------------------------------------------------
$sPesquisaPagamento[$eClickPaginarTLoteX]	=	1315
$sPesquisaPagamento[$eClickPaginarTLoteY]	=	577
;------------------------------------------------------

;------------------------------------------------------
;Ajustar de forma que não pegue o caractere especial (*) no print
$sPesquisaPagamento[$eMapaPrintTLoteInicioX]	=	50
$sPesquisaPagamento[$eMapaPrintTLoteInicioY]	=	442
$sPesquisaPagamento[$eMapaPrintTLoteFimX]		=	395
$sPesquisaPagamento[$eMapaPrintTLoteFimY]		=	606
;------------------------------------------------------


;------------------------------------------------------
$sPesquisaPagamento[$eFecharTelaPgtoX]	= 1350
$sPesquisaPagamento[$eFecharTelaPgtoY]	= 96
;-------------------------------------------------------

;------------------------------------------------------
$sReversaoPagamento[$eOpcaoPesquisaPgtoX]	=	240
$sReversaoPagamento[$eOpcaoPesquisaPgtoY]	=	65
;------------------------------------------------------

;------------------------------------------------------
$sReversaoPagamento[$eMaximizarTelaPgtoX]	=	535
$sReversaoPagamento[$eMaximizarTelaPgtoY]	=	100
;------------------------------------------------------

;------------------------------------------------------
$sReversaoPagamento[$eColorGridTLoteX]	=	1316
$sReversaoPagamento[$eColorGridTLoteY]	=	579
;------------------------------------------------------

;------------------------------------------------------
$sReversaoPagamento[$eClickPaginarTLoteX]	=	1315
$sReversaoPagamento[$eClickPaginarTLoteY]	=	577
;------------------------------------------------------

;------------------------------------------------------
$sReversaoPagamento[$eFecharTelaPgtoX]	= 1350
$sReversaoPagamento[$eFecharTelaPgtoY]	= 96
;-------------------------------------------------------

Global $sApp[11]
Global Enum $eAutomatioName,$eSleepTime,$USER,$PWS,$oIE,$url,$formID,$formUID,$formPID,$uName,$pwd;,$oForm,$o_txt_user,$o_txt_pwd,$hFile,$s_Directory,$data

;~ ### CFG ROBO
	$sApp[$eSleepTime]		=	2000
	$sApp[$uName]			=	""


Global $arrTela[8]
Global Enum $eWinLogon,$eAtlys,$eNgin,$eVivoNet,$eAtlysSolution,$eBloqueioAtlys,$eCSO,$eAtlysErro
	$arrTela[$eAtlys] 			= "Atlys Global Solution - \\Remote"
	$arrTela[$eAtlysSolution] 	= "Logon do Atlys Global Solution - \\Remote"
	$arrTela[$eBloqueioAtlys] 	= "Bloqueio da Solução Global Atlys - \\Remote"
	$arrTela[$eAtlysErro]		= "Atlys ERROR"

Global $arrLote[3]
Global Enum $eLote,$eContaCritica,$eContaRevisao
	$arrLote[$eLote] 			= "329975"
	$arrLote[$eContaCritica] 	= "6400001529"
	$arrLote[$eContaRevisao] 	= "0222046478"





Global $arrayS1ListaPosicao[400]
Global Enum $eS4T0MsgSenhaErradaInicioX, _
			$eS4T0MsgSenhaErradaInicioY, _
			$eS4T0MsgSenhaErradaFimX, _
			$eS4T0MsgSenhaErradaFimY, _
			$eS1T0MsgSenhaErradaInicioX, _
			$eS1T0MsgSenhaErradaInicioY, _
			$eS1T0MsgSenhaErradaFimX, _
			$eS1T0MsgSenhaErradaFimY, _
			$eS1T1BotaoFecharTelaPrincipalX, _
			$eS1T1BotaoFecharTelaPrincipalY, _
			$eS1T1BotaoMaximizarX, _
			$eS1T1BotaoMaximizarY, _
			$eS1T1PesquisarInicioX, _
			$eS1T1PesquisarInicioY, _
			$eS1T1PesquisarFimX, _
			$eS1T1PesquisarFimY, _
			$eS1T1ComboPesquisarPorX, _
			$eS1T1ComboPesquisarPorY, _
			$eS1T1ComboSelecionarIdentificacaoClienteX, _
			$eS1T1ComboSelecionarIdentificacaoClienteY, _
			$eS1T1ClicarBotaoPesquisarX, _
			$eS1T1ClicarBotaoPesquisarY, _
			$eS1T1PrintCapaCnpjInicioX, _
			$eS1T1PrintCapaCnpjInicioY, _
			$eS1T1PrintCapaCnpjFimX, _
			$eS1T1PrintCapaCnpjFimY, _
			$eS1T1PaginacaoCnpjCapaX, _
			$eS1T1PaginacaoCnpjCapaY, _
			$eS1T1PaginacaoColorX, _
			$eS1T1PaginacaoColorY, _
			$eS1T2TelaResumoCobrancaFecharInicioX, _
			$eS1T2TelaResumoCobrancaFecharInicioY, _
			$eS1T2TelaResumoCobrancaFecharFimX, _
			$eS1T2TelaResumoCobrancaFecharFimY, _
			$eS1T2TelaResumoCobrancaSelecionarOpcaoContaX, _
			$eS1T2TelaResumoCobrancaSelecionarOpcaoContaY, _
			$eS1T2TelaResumoCobrancaMaximizarX, _
			$eS1T2TelaResumoCobrancaMaximizarY, _
			$eS1T2TelaResumoCobrancaInicioX, _
			$eS1T2TelaResumoCobrancaInicioY, _
			$eS1T2TelaResumoCobrancaFimX, _
			$eS1T2TelaResumoCobrancaFimY, _
			$eS1T2TelaResumoCobrancaClickPagX, _
			$eS1T2TelaResumoCobrancaClickPagY, _
			$eS1T2TelaResumoCobrancaMinizarTelaX, _
			$eS1T2TelaResumoCobrancaMinizarTelaY, _
			$eS1T2TelaResumoCobrancaColorPixelPagX, _
			$eS1T2TelaResumoCobrancaColorPixelPagY, _
			$eS1T2TelaResumoCobrancaFecharX, _
			$eS1T2TelaResumoCobrancaFecharY, _
			$eS1T1PesquisarPorContaX, _
			$eS1T1PesquisarPorContaY, _
			$eS1T1ClickVisualizarX, _
			$eS1T1ClickVisualizarY, _
			$eS1T3IdentTelaNotasBotaoFecharInicioX, _
			$eS1T3IdentTelaNotasBotaoFecharInicioY, _
			$eS1T3IdentTelaNotasBotaoFecharFimX, _
			$eS1T3IdentTelaNotasBotaoFecharFimY, _
			$eS1T3IdentTelaNotasBotaoNotasInicioX, _
			$eS1T3IdentTelaNotasBotaoNotasInicioY, _
			$eS1T3IdentTelaNotasBotaoNotasFimX, _
			$eS1T3IdentTelaNotasBotaoNotasFimY, _
			$eS1T3ClicarBotaNotasX, _
			$eS1T3ClicarBotaNotasY, _
			$eS1T3ClicarOpcaoCrescenteX, _
			$eS1T3ClicarOpcaoCrescenteY, _
			$eS1T3ClicarOpcaoTudoX, _
			$eS1T3ClicarOpcaoTudoY, _
			$eS1T3ClicarBotaoPesquisarX, _
			$eS1T3ClicarBotaoPesquisarY, _
			$eS1T3ClicarTelaNotasPixelColorX, _
			$eS1T3ClicarTelaNotasPixelColorY, _
			$eS1T3ClicarTelaNotasPaginacaoX, _
			$eS1T3ClicarTelaNotasPaginacaoY, _
			$eS1T3TelaNotasPrintInicioX, _
			$eS1T3TelaNotasPrintInicioY, _
			$eS1T3TelaNotasPrintFimX, _
			$eS1T3TelaNotasPrintFimY, _
			$eS1T3TelaNotasBotaoFecharX, _
			$eS1T3TelaNotasBotaoFecharY, _
			$eS1T4TelaCapaContaPrintInicioX, _
			$eS1T4TelaCapaContaPrintInicioY, _
			$eS1T4TelaCapaContaPrintFimX, _
			$eS1T4TelaCapaContaPrintFimY, _
			$eS1T4TelaCapaContaIdentBtnAssinaturaInicioX, _
			$eS1T4TelaCapaContaIdentBtnAssinaturaInicioY, _
			$eS1T4TelaCapaContaIdentBtnAssinaturaFimX, _
			$eS1T4TelaCapaContaIdentBtnAssinaturaFimY, _
			$eS1T4TelaCapaContaSelecionaTelaX, _
			$eS1T4TelaCapaContaSelecionaTelaY, _
			$eS1T4TelaCapaContaClickBotaoAssinaturaX, _
			$eS1T4TelaCapaContaClickBotaoAssinaturaY, _
			$eS1T5TelaIdentBtnAdicionarInicioX, _
			$eS1T5TelaIdentBtnAdicionarInicioY, _
			$eS1T5TelaIdentBtnAdicionarFimX, _
			$eS1T5TelaIdentBtnAdicionarFimY, _
			$eS1T5TelaIdentBtnFecharInicioX, _
			$eS1T5TelaIdentBtnFecharInicioY, _
			$eS1T5TelaIdentBtnFecharFimX, _
			$eS1T5TelaIdentBtnFecharFimY, _
			$eS1T5TelaIdentPixelColorX, _
			$eS1T5TelaIdentPixelColorY, _
			$eS1T5TelaListaTerminaisPrintInicioX , _
			$eS1T5TelaListaTerminaisPrintInicioY , _
			$eS1T5TelaListaTerminaisPrintFimX , _
			$eS1T5TelaListaTerminaisPrintFimY, _
			$eS1T5TelaListaTerminaisPaginacaoX, _
			$eS1T5TelaListaTerminaisPaginacaoY, _
			$eS1T6PgtoAjusteIdentBotaoPagamentoInicioX , _
			$eS1T6PgtoAjusteIdentBotaoPagamentoInicioY , _
			$eS1T6PgtoAjusteIdentBotaoPagamentoFimX , _
			$eS1T6PgtoAjusteIdentBotaoPagamentoFimY, _
			$eS1T6PgtoAjusteClickBotaoPagamentoX , _
			$eS1T6PgtoAjusteClickBotaoPagamentoY , _
			$eS1T6PgtoAjusteIdentBotaoPesquisarInicioX , _
			$eS1T6PgtoAjusteIdentBotaoPesquisarInicioY , _
			$eS1T6PgtoAjusteIdentBotaoPesquisarFimX , _
			$eS1T6PgtoAjusteIdentBotaoPesquisarFimY , _
			$eS1T6PgtoAjusteClickMaximizarX , _
			$eS1T6PgtoAjusteClickMaximizarY , _
			$eS1T6PgtoAjusteCapturaPixelColorX, _
			$eS1T6PgtoAjusteCapturaPixelColorY, _
			$eS1T6PgtoAjusteClickPaginacaoX, _
			$eS1T6PgtoAjusteClickPaginacaoY, _
			$eS1T6PgtoAjustePrintTelaInicioX , _
			$eS1T6PgtoAjustePrintTelaInicioY , _
			$eS1T6PgtoAjustePrintTelaFimX , _
			$eS1T6PgtoAjustePrintTelaFimY, _
			$eS1T6PgtoAjusteClickOpcaoAjusteX, _
			$eS1T6PgtoAjusteClickOpcaoAjusteY, _
			$eS1T6PgtoAjusteCapturaPixelColorOpcaoAjusteX, _
			$eS1T6PgtoAjusteCapturaPixelColorOpcaoAjusteY, _
			$eS1T6PgtoAjusteClickPaginacaoOpcaoAjusteX, _
			$eS1T6PgtoAjusteClickPaginacaoOpcaoAjusteY, _
			$eS1T6PgtoAjusteMinimizarTelaX, _
			$eS1T6PgtoAjusteMinimizarTelaY, _
			$eS1T6PgtoAjusteFecharX, _
			$eS1T6PgtoAjusteFecharY, _
			$eS1T7PrintCapaTerminalInicioX, _
			$eS1T7PrintCapaTerminalInicioY, _
			$eS1T7PrintCapaTerminalFimX, _
			$eS1T7PrintCapaTerminalFimY, _
			$eS1T7CapaTerminalBotaoLimparX, _
			$eS1T7CapaTerminalBotaoLimparY, _
			$eS1T7CapaTerminalBotaoVisualizarX, _
			$eS1T7CapaTerminalBotaoVisualizarY, _
			$eS1T8TerminalIdentificarFuncAssinaturaInicioX, _
			$eS1T8TerminalIdentificarFuncAssinaturaInicioY, _
			$eS1T8TerminalIdentificarFuncAssinaturaFimX, _
			$eS1T8TerminalIdentificarFuncAssinaturaFimY, _
			$eS1T8TerminalFuncAssinaturaClickX, _
			$eS1T8TerminalFuncAssinaturaClickY, _
			$eS1T8TerminalDescMktClickX, _
			$eS1T8TerminalDescMktClickY, _
			$eS1T9BotaoAdicionarInicioX, _
			$eS1T9BotaoAdicionarInicioY, _
			$eS1T9BotaoAdicionarFimX, _
			$eS1T9BotaoAdicionarFimY, _
			$eS1T9PrintMktInicioX, _
			$eS1T9PrintMktInicioY, _
			$eS1T9PrintMktFimX, _
			$eS1T9PrintMktFimY, _
			$eS1T9MktBotaoFecharX, _
			$eS1T9MktBotaoFecharY, _
			$eS1T10IdentBotaoSalvarInicioX, _
			$eS1T10IdentBotaoSalvarInicioY, _
			$eS1T10IdentBotaoSalvarFimX, _
			$eS1T10IdentBotaoSalvarFimY, _
			$eS1T10GetColorPixelX, _
			$eS1T10GetColorPixelY, _
			$eS1T10HistoricoPrintInicioX, _
			$eS1T10HistoricoPrintInicioY, _
			$eS1T10HistoricoPrintFimX, _
			$eS1T10HistoricoPrintFimY, _
			$eS1T10HistoricoPaginacaoLateralDireitaX, _
			$eS1T10HistoricoPaginacaoLateralDireitaY, _
			$eS1T10HistoricoPaginacaoLateralEsquerdaX, _
			$eS1T10HistoricoPaginacaoLateralEsquerdaY, _
			$eS1T10HistoricoPaginacaoX, _
			$eS1T10HistoricoPaginacaoY, _
			$eS1T10ClickMinimizarX, _
			$eS1T10ClickMinimizarY, _
			$eS1T10ClickBotaoFecharX, _
			$eS1T10ClickBotaoFecharY, _
			$eS1T10OpcaoDescontoMktClickX, _
			$eS1T10OpcaoDescontoMktClickY, _
			$eS1T10OpcaoHistoricoClickX, _
			$eS1T10OpcaoHistoricoClickY, _
			$eS1T10ClickMaximizarX, _
			$eS1T10ClickMaximizarY, _
			$eS1T10TelaCapaTerminalClickBotaoFecharX, _
			$eS1T10TelaCapaTerminalClickBotaoFecharY, _
			$eS1T9SelecionarCapaTerminalX, _
			$eS1T9SelecionarCapaTerminalY, _
			$eS1T10TelaTerminaisSelecionarX, _
			$eS1T10TelaTerminaisSelecionarY, _
			$eS1T7SelecionarTerminalX, _
			$eS1T7SelecionarTerminalY, _
			$eS1T10MensagemPesquisaNaoEncontradaInicioX,  _
			$eS1T10MensagemPesquisaNaoEncontradaInicioY,  _
			$eS1T10MensagemPesquisaNaoEncontradaFimX,  _
			$eS1T10MensagemPesquisaNaoEncontradaFimY, _
			$eS1T7CapaTerminalBotaoPesquisarX, _
			$eS1T7CapaTerminalBotaoPesquisarY, _
			$eS1T5TelaListaTerminaisBotaoFecharX, _
			$eS1T5TelaListaTerminaisBotaoFecharY, _
			$eS1T5TelaCapaContaBotaoFecharX, _
			$eS1T5TelaCapaContaBotaoFecharY, _
			$eS1T5TelaAssinaturasBotaoFecharX, _
			$eS1T5TelaAssinaturasBotaoFecharY, _
			$eS1T1PesquisarPorContaPaginacaoX, _
			$eS1T1PesquisarPorContaPaginacaoY, _
			$eS1T2ValidaRetornoPesquisaInicioX, _
			$eS1T2ValidaRetornoPesquisaInicioY, _
			$eS1T2ValidaRetornoPesquisaFimX, _
			$eS1T2ValidaRetornoPesquisaFimY, _
			$eS1T1BotaoAssinaturaInicioX, _
			$eS1T1BotaoAssinaturaInicioY, _
			$eS1T1BotaoAssinaturaFimX, _
			$eS1T1BotaoAssinaturaFimY






;Verificar Mensagem de Senha Errada Citrix
;===============================================
$arrayS1ListaPosicao[$eS4T0MsgSenhaErradaInicioX] = 597
$arrayS1ListaPosicao[$eS4T0MsgSenhaErradaInicioY] = 432
$arrayS1ListaPosicao[$eS4T0MsgSenhaErradaFimX] = 790
$arrayS1ListaPosicao[$eS4T0MsgSenhaErradaFimY] = 462
;===============================================

;Verificar Mensagem de Senha Errada Atlys
;===============================================
$arrayS1ListaPosicao[$eS1T0MsgSenhaErradaInicioX] = 500
$arrayS1ListaPosicao[$eS1T0MsgSenhaErradaInicioY] = 371
$arrayS1ListaPosicao[$eS1T0MsgSenhaErradaFimX] = 893
$arrayS1ListaPosicao[$eS1T0MsgSenhaErradaFimY] = 417
;===============================================

;Maximiar Tela Inicial Por Click Atlys
;===============================================
$arrayS1ListaPosicao[$eS1T1BotaoMaximizarX] = 738
$arrayS1ListaPosicao[$eS1T1BotaoMaximizarY] = 11
;===============================================

;Verificar Texto Pesquisar Atlys
;===============================================
$arrayS1ListaPosicao[$eS1T1PesquisarInicioX] = 15
$arrayS1ListaPosicao[$eS1T1PesquisarInicioY] = 134
$arrayS1ListaPosicao[$eS1T1PesquisarFimX] = 67
$arrayS1ListaPosicao[$eS1T1PesquisarFimY] = 156
;===============================================

;Selecionar Combo Pesquisar Por (Click)
;===============================================
$arrayS1ListaPosicao[$eS1T1ComboPesquisarPorX] = 224
$arrayS1ListaPosicao[$eS1T1ComboPesquisarPorY] = 148
;===============================================

;Selecionar Combo Selecionar Identificacao Cliente (Click)
;====================================================================
$arrayS1ListaPosicao[$eS1T1ComboSelecionarIdentificacaoClienteX] = 157
$arrayS1ListaPosicao[$eS1T1ComboSelecionarIdentificacaoClienteY] = 167
;====================================================================

;Clicar no Botao Pesquisar Apos Preeenchimento campo Text
;====================================================================
$arrayS1ListaPosicao[$eS1T1ClicarBotaoPesquisarX] = 65
$arrayS1ListaPosicao[$eS1T1ClicarBotaoPesquisarY] = 279
;====================================================================

;Posicao Prin Tela Capa Cnpj
;====================================================================
$arrayS1ListaPosicao[$eS1T1PrintCapaCnpjInicioX]  = 0
$arrayS1ListaPosicao[$eS1T1PrintCapaCnpjInicioY]  = 84
$arrayS1ListaPosicao[$eS1T1PrintCapaCnpjFimX]		= 698
$arrayS1ListaPosicao[$eS1T1PrintCapaCnpjFimY] 	= 574
;====================================================================

;Posicao Paginacao Tela Cnpj Click
;====================================================================
$arrayS1ListaPosicao[$eS1T1PaginacaoCnpjCapaX]	= 668
$arrayS1ListaPosicao[$eS1T1PaginacaoCnpjCapaY]	= 478
;====================================================================

;Posicao Paginacao Tela Cnpj Click
;====================================================================
$arrayS1ListaPosicao[$eS1T1PaginacaoColorX]	= 668
$arrayS1ListaPosicao[$eS1T1PaginacaoColorY]	= 467
;====================================================================

;Posicao Botao Fechar Tela Resumo Cobranca
;====================================================================
$arrayS1ListaPosicao[$eS1T2TelaResumoCobrancaFecharInicioX] = 785
$arrayS1ListaPosicao[$eS1T2TelaResumoCobrancaFecharInicioY] = 622
$arrayS1ListaPosicao[$eS1T2TelaResumoCobrancaFecharFimX] 	  =	860
$arrayS1ListaPosicao[$eS1T2TelaResumoCobrancaFecharFimY]	  = 644
;====================================================================

;Posicao Selecionar Opcao Conta
;=========================================================================
$arrayS1ListaPosicao[$eS1T2TelaResumoCobrancaSelecionarOpcaoContaX]	= 283
$arrayS1ListaPosicao[$eS1T2TelaResumoCobrancaSelecionarOpcaoContaY]	= 177
;=========================================================================

;Maximizar Tela Resumo Cobranca
;=========================================================================
$arrayS1ListaPosicao[$eS1T2TelaResumoCobrancaMaximizarX]	= 850
$arrayS1ListaPosicao[$eS1T2TelaResumoCobrancaMaximizarY]	= 120
;=========================================================================

;Tela Resumo Cobranca Paginacao
;=========================================================================
$arrayS1ListaPosicao[$eS1T2TelaResumoCobrancaClickPagX]	= 1317
$arrayS1ListaPosicao[$eS1T2TelaResumoCobrancaClickPagY]	= 593
;=========================================================================

;Posicao Tela Resumo Cobranca Print
;=========================================================================
$arrayS1ListaPosicao[$eS1T2TelaResumoCobrancaInicioX] = 0
$arrayS1ListaPosicao[$eS1T2TelaResumoCobrancaInicioY] = 83
$arrayS1ListaPosicao[$eS1T2TelaResumoCobrancaFimX] 	= 1365
$arrayS1ListaPosicao[$eS1T2TelaResumoCobrancaFimY]	= 749
;=========================================================================

;Click Paginacao Resumo Cobranca Busca Pixel
;=========================================================================
$arrayS1ListaPosicao[$eS1T2TelaResumoCobrancaColorPixelPagX]	= 1318
$arrayS1ListaPosicao[$eS1T2TelaResumoCobrancaColorPixelPagY]	= 586
;=========================================================================

;Click Resumo Cobranca Botao Minimizar Tela
;=========================================================================
$arrayS1ListaPosicao[$eS1T2TelaResumoCobrancaMinizarTelaX]	= 1332
$arrayS1ListaPosicao[$eS1T2TelaResumoCobrancaMinizarTelaY]	= 95
;=========================================================================

;Click Resumo Cobranca Botao Fechar
;=========================================================================
$arrayS1ListaPosicao[$eS1T2TelaResumoCobrancaFecharX]	= 821
$arrayS1ListaPosicao[$eS1T2TelaResumoCobrancaFecharY]	= 634
;=========================================================================

;Clicar Visualizar/Alterar Conta
;=========================================================================
$arrayS1ListaPosicao[$eS1T1ClickVisualizarX]	= 200
$arrayS1ListaPosicao[$eS1T1ClickVisualizarY]	= 505
;=========================================================================

;Identificar Tela Notas de Contas - Botao Fechar
;=========================================================================
$arrayS1ListaPosicao[$eS1T3IdentTelaNotasBotaoFecharInicioX]	= 674
$arrayS1ListaPosicao[$eS1T3IdentTelaNotasBotaoFecharInicioY]	= 490
$arrayS1ListaPosicao[$eS1T3IdentTelaNotasBotaoFecharFimX]		= 774
$arrayS1ListaPosicao[$eS1T3IdentTelaNotasBotaoFecharFimY]		= 512
;=========================================================================

;Identificar Tela Notas de Contas - Botao Notas
;=========================================================================
$arrayS1ListaPosicao[$eS1T3IdentTelaNotasBotaoNotasInicioX]	= 658
$arrayS1ListaPosicao[$eS1T3IdentTelaNotasBotaoNotasInicioY]	= 638
$arrayS1ListaPosicao[$eS1T3IdentTelaNotasBotaoNotasFimX]		= 759
$arrayS1ListaPosicao[$eS1T3IdentTelaNotasBotaoNotasFimY]		= 661
;=========================================================================

;Clicar Botao Notas
;=========================================================================
$arrayS1ListaPosicao[$eS1T3ClicarBotaNotasX]	= 708
$arrayS1ListaPosicao[$eS1T3ClicarBotaNotasY]	= 648
;=========================================================================

;Clicar Opcao Crescente
;=========================================================================
$arrayS1ListaPosicao[$eS1T3ClicarOpcaoCrescenteX]	= 639
$arrayS1ListaPosicao[$eS1T3ClicarOpcaoCrescenteY]	= 148
;=========================================================================

;Clicar Opcao Pesquisar
;=========================================================================
$arrayS1ListaPosicao[$eS1T3ClicarOpcaoTudoX]	= 39
$arrayS1ListaPosicao[$eS1T3ClicarOpcaoTudoY]	= 155
;=========================================================================

;Clicar Opcao Pesquisar
;=========================================================================
$arrayS1ListaPosicao[$eS1T3ClicarBotaoPesquisarX]	= 75
$arrayS1ListaPosicao[$eS1T3ClicarBotaoPesquisarY]	= 184
;=========================================================================

;Identificar Hexadecimal Cor Pixel Tela Notas
;=========================================================================
$arrayS1ListaPosicao[$eS1T3ClicarTelaNotasPixelColorX]	= 756
$arrayS1ListaPosicao[$eS1T3ClicarTelaNotasPixelColorY]	= 423
;=========================================================================

;Paginacao Tela Notas Contas
;=========================================================================
$arrayS1ListaPosicao[$eS1T3ClicarTelaNotasPaginacaoX]	= 756
$arrayS1ListaPosicao[$eS1T3ClicarTelaNotasPaginacaoY]	= 435
;=========================================================================

;Tela Notas Contas Print
;=========================================================================
$arrayS1ListaPosicao[$eS1T3TelaNotasPrintInicioX]	= 10
$arrayS1ListaPosicao[$eS1T3TelaNotasPrintInicioY]	= 7
$arrayS1ListaPosicao[$eS1T3TelaNotasPrintFimX]	= 788
$arrayS1ListaPosicao[$eS1T3TelaNotasPrintFimY]	= 528
;=========================================================================

;Clicar Botao Fechar Notas Contas
;=========================================================================
$arrayS1ListaPosicao[$eS1T3TelaNotasBotaoFecharX]	= 724
$arrayS1ListaPosicao[$eS1T3TelaNotasBotaoFecharY]	= 500
;=========================================================================

;Selecionar Opcao Pesquisar Por Conta
;=========================================================================
$arrayS1ListaPosicao[$eS1T1PesquisarPorContaX]	= 139
$arrayS1ListaPosicao[$eS1T1PesquisarPorContaY]	= 269
;=========================================================================

;Selecionar Opcao Pesquisar Por Conta Paginacao
;=========================================================================
$arrayS1ListaPosicao[$eS1T1PesquisarPorContaPaginacaoX]	= 228
$arrayS1ListaPosicao[$eS1T1PesquisarPorContaPaginacaoY]	= 284
;=========================================================================

;Tela Capa Conta - Print
;=========================================================================
$arrayS1ListaPosicao[$eS1T4TelaCapaContaPrintInicioX] = 14
$arrayS1ListaPosicao[$eS1T4TelaCapaContaPrintInicioY] = 105
$arrayS1ListaPosicao[$eS1T4TelaCapaContaPrintFimX]	= 774
$arrayS1ListaPosicao[$eS1T4TelaCapaContaPrintFimY]	= 717
;==========================================================================

;Tela Capa Conta Identificar Botao Assinatura
;=========================================================================
$arrayS1ListaPosicao[$eS1T4TelaCapaContaIdentBtnAssinaturaInicioX] = 349
$arrayS1ListaPosicao[$eS1T4TelaCapaContaIdentBtnAssinaturaInicioY] = 672
$arrayS1ListaPosicao[$eS1T4TelaCapaContaIdentBtnAssinaturaFimX]	 = 453
$arrayS1ListaPosicao[$eS1T4TelaCapaContaIdentBtnAssinaturaFimY]	 = 700
;==========================================================================

;Tela Capa Conta Selecionar Tela
;=========================================================================
$arrayS1ListaPosicao[$eS1T4TelaCapaContaSelecionaTelaX] = 254
$arrayS1ListaPosicao[$eS1T4TelaCapaContaSelecionaTelaY] = 683
;=========================================================================

;Tela Capa Conta Click Botao Assinatura
;=========================================================================
$arrayS1ListaPosicao[$eS1T4TelaCapaContaClickBotaoAssinaturaX] = 421
$arrayS1ListaPosicao[$eS1T4TelaCapaContaClickBotaoAssinaturaY] = 684
;=========================================================================

;Tela Pesquisa Por Acesso
;=========================================================================
$arrayS1ListaPosicao[$eS1T5TelaIdentBtnAdicionarInicioX] 	= 37
$arrayS1ListaPosicao[$eS1T5TelaIdentBtnAdicionarInicioY] 	= 642
$arrayS1ListaPosicao[$eS1T5TelaIdentBtnAdicionarFimX]	 	= 165
$arrayS1ListaPosicao[$eS1T5TelaIdentBtnAdicionarFimY]	 	= 673
;==========================================================================

;Identificar Pixel Color Tela List Terminais
;=========================================================================
$arrayS1ListaPosicao[$eS1T5TelaIdentPixelColorX] = 768
$arrayS1ListaPosicao[$eS1T5TelaIdentPixelColorY] = 525
;=========================================================================

;Print Tela Lista Terminais
;=========================================================================
$arrayS1ListaPosicao[$eS1T5TelaListaTerminaisPrintInicioX] = 28
$arrayS1ListaPosicao[$eS1T5TelaListaTerminaisPrintInicioY] = 128
$arrayS1ListaPosicao[$eS1T5TelaListaTerminaisPrintFimX] 	 = 794
$arrayS1ListaPosicao[$eS1T5TelaListaTerminaisPrintFimY] 	 = 679
;=========================================================================

;Tela Lista Terminais Paginacao
;=========================================================================
$arrayS1ListaPosicao[$eS1T5TelaListaTerminaisPaginacaoX] = 768
$arrayS1ListaPosicao[$eS1T5TelaListaTerminaisPaginacaoY] = 536
;=========================================================================

;Tela Identificar Botao Pagamento
;=========================================================================
$arrayS1ListaPosicao[$eS1T6PgtoAjusteIdentBotaoPagamentoInicioX] 	= 358
$arrayS1ListaPosicao[$eS1T6PgtoAjusteIdentBotaoPagamentoInicioY] 	= 597
$arrayS1ListaPosicao[$eS1T6PgtoAjusteIdentBotaoPagamentoFimX]    	= 465
$arrayS1ListaPosicao[$eS1T6PgtoAjusteIdentBotaoPagamentoFimY]		= 619
;=========================================================================

;Botao Pagamento Click item 1
;=========================================================================
$arrayS1ListaPosicao[$eS1T6PgtoAjusteClickBotaoPagamentoX]	= 420
$arrayS1ListaPosicao[$eS1T6PgtoAjusteClickBotaoPagamentoY]	= 610
;=========================================================================

;Tela Identificar Pgto e Ajuste Botao Pesquisar
;=========================================================================
$arrayS1ListaPosicao[$eS1T6PgtoAjusteIdentBotaoPesquisarInicioX] 	= 44
$arrayS1ListaPosicao[$eS1T6PgtoAjusteIdentBotaoPesquisarInicioY] 	= 282
$arrayS1ListaPosicao[$eS1T6PgtoAjusteIdentBotaoPesquisarFimX]    	= 130
$arrayS1ListaPosicao[$eS1T6PgtoAjusteIdentBotaoPesquisarFimY]		= 303
;=========================================================================

;Tela Pagto Ajuste Click Maximizar
;=========================================================================
$arrayS1ListaPosicao[$eS1T6PgtoAjusteClickMaximizarX] 	= 804
$arrayS1ListaPosicao[$eS1T6PgtoAjusteClickMaximizarY] 	= 136
;=========================================================================

;Tela Pagto Busca Color Pixel
;=========================================================================
$arrayS1ListaPosicao[$eS1T6PgtoAjusteCapturaPixelColorX] 	= 1307
$arrayS1ListaPosicao[$eS1T6PgtoAjusteCapturaPixelColorY] 	= 546
;=========================================================================

;Tela Pagto Paginacao
;=========================================================================
$arrayS1ListaPosicao[$eS1T6PgtoAjusteClickPaginacaoX] 	= 1307
$arrayS1ListaPosicao[$eS1T6PgtoAjusteClickPaginacaoY] 	= 558
;=========================================================================

;Tela Pgto e Ajuste Print
;=========================================================================
$arrayS1ListaPosicao[$eS1T6PgtoAjustePrintTelaInicioX] 	= 0
$arrayS1ListaPosicao[$eS1T6PgtoAjustePrintTelaInicioY] 	= 80
$arrayS1ListaPosicao[$eS1T6PgtoAjustePrintTelaFimX]    	= 1365
$arrayS1ListaPosicao[$eS1T6PgtoAjustePrintTelaFimY]		= 753
;=========================================================================

;Tela Pagto Click Ajuste
;=========================================================================
$arrayS1ListaPosicao[$eS1T6PgtoAjusteClickOpcaoAjusteX] 	= 193
$arrayS1ListaPosicao[$eS1T6PgtoAjusteClickOpcaoAjusteY] 	= 325
;=========================================================================


;Tela Pagto Busca Color Pixel Opcao Ajuste
;=========================================================================
$arrayS1ListaPosicao[$eS1T6PgtoAjusteCapturaPixelColorOpcaoAjusteX] 	= 1307
$arrayS1ListaPosicao[$eS1T6PgtoAjusteCapturaPixelColorOpcaoAjusteY] 	= 524
;=========================================================================

;Tela Pagto Paginacao
;=========================================================================
$arrayS1ListaPosicao[$eS1T6PgtoAjusteClickPaginacaoOpcaoAjusteX] 	= 1307
$arrayS1ListaPosicao[$eS1T6PgtoAjusteClickPaginacaoOpcaoAjusteY] 	= 538
;=========================================================================

;Tela Pagto Minimizar
;=========================================================================
$arrayS1ListaPosicao[$eS1T6PgtoAjusteMinimizarTelaX] 	= 1334
$arrayS1ListaPosicao[$eS1T6PgtoAjusteMinimizarTelaY] 	= 94
;=========================================================================

;Tela Pagto Fechar
;=========================================================================
$arrayS1ListaPosicao[$eS1T6PgtoAjusteFecharX] 	= 747
$arrayS1ListaPosicao[$eS1T6PgtoAjusteFecharY] 	= 677
;=========================================================================

;Tela Capa Terminal
;=========================================================================
$arrayS1ListaPosicao[$eS1T7PrintCapaTerminalInicioX] 	= 26
$arrayS1ListaPosicao[$eS1T7PrintCapaTerminalInicioY] 	= 126
$arrayS1ListaPosicao[$eS1T7PrintCapaTerminalFimX]    	= 796
$arrayS1ListaPosicao[$eS1T7PrintCapaTerminalFimY]		= 676
;=========================================================================

;Tela Terminal Click Visualizar
;=========================================================================
$arrayS1ListaPosicao[$eS1T7CapaTerminalBotaoVisualizarX] 	= 94
$arrayS1ListaPosicao[$eS1T7CapaTerminalBotaoVisualizarY] 	= 582
;=========================================================================

;Tela Terminal Identificar Func Assinatura
;=========================================================================
$arrayS1ListaPosicao[$eS1T8TerminalIdentificarFuncAssinaturaInicioX] 	= 630
$arrayS1ListaPosicao[$eS1T8TerminalIdentificarFuncAssinaturaInicioY]	= 551
$arrayS1ListaPosicao[$eS1T8TerminalIdentificarFuncAssinaturaFimX]		= 821
$arrayS1ListaPosicao[$eS1T8TerminalIdentificarFuncAssinaturaFimY]		= 574
;=========================================================================

;Tela Terminal Botao Func Assinatura Click
;=========================================================================
$arrayS1ListaPosicao[$eS1T8TerminalFuncAssinaturaClickX] 	= 538
$arrayS1ListaPosicao[$eS1T8TerminalFuncAssinaturaClickY] 	= 562
;========================================================================

;Tela Desc Mkt Identificar Botao Adicionar
;=========================================================================
$arrayS1ListaPosicao[$eS1T9BotaoAdicionarInicioX] 	= 31
$arrayS1ListaPosicao[$eS1T9BotaoAdicionarInicioY] 	= 341
$arrayS1ListaPosicao[$eS1T9BotaoAdicionarFimX] 		= 99
$arrayS1ListaPosicao[$eS1T9BotaoAdicionarFimY] 		= 364
;========================================================================

;Tela Desc Mkt Print
;=========================================================================
$arrayS1ListaPosicao[$eS1T9PrintMktInicioX] 	= 20
$arrayS1ListaPosicao[$eS1T9PrintMktInicioY]	= 16
$arrayS1ListaPosicao[$eS1T9PrintMktFimX]		= 612
$arrayS1ListaPosicao[$eS1T9PrintMktFimY]		= 376
;=========================================================================

;Tela Historico Identificar Botao Salvar
;=========================================================================
$arrayS1ListaPosicao[$eS1T10IdentBotaoSalvarInicioX] 	= 227
$arrayS1ListaPosicao[$eS1T10IdentBotaoSalvarInicioY]	= 470
$arrayS1ListaPosicao[$eS1T10IdentBotaoSalvarFimX]		= 358
$arrayS1ListaPosicao[$eS1T10IdentBotaoSalvarFimY]		= 495
;=========================================================================

;Tela Historico Click Maximizar
;=========================================================================
$arrayS1ListaPosicao[$eS1T10ClickMaximizarX] 	= 1028
$arrayS1ListaPosicao[$eS1T10ClickMaximizarY] 	= 172
;========================================================================

;Tela Historico Identificar Color Pixel
;=========================================================================
$arrayS1ListaPosicao[$eS1T10GetColorPixelX] 	= 1322
$arrayS1ListaPosicao[$eS1T10GetColorPixelY] 	= 565
;========================================================================

;Tela Print Historico
;=========================================================================
$arrayS1ListaPosicao[$eS1T10HistoricoPrintInicioX] 	= 0
$arrayS1ListaPosicao[$eS1T10HistoricoPrintInicioY]	= 84
$arrayS1ListaPosicao[$eS1T10HistoricoPrintFimX]		= 1365
$arrayS1ListaPosicao[$eS1T10HistoricoPrintFimY]		= 748
;=========================================================================

;Tela Historico Paginacao Lateral Direita
;=========================================================================
$arrayS1ListaPosicao[$eS1T10HistoricoPaginacaoLateralDireitaX] 	= 1305
$arrayS1ListaPosicao[$eS1T10HistoricoPaginacaoLateralDireitaY] 	= 595
;========================================================================

;Tela Historico Paginacao Vertical
;=========================================================================
$arrayS1ListaPosicao[$eS1T10HistoricoPaginacaoX] 	= 1324
$arrayS1ListaPosicao[$eS1T10HistoricoPaginacaoY] 	= 577
;========================================================================

;Tela Historico Click Minimizar
;=========================================================================
$arrayS1ListaPosicao[$eS1T10ClickMinimizarX] 	= 1333
$arrayS1ListaPosicao[$eS1T10ClickMinimizarY] 	= 96
;========================================================================

;Tela Historico Botao Fechar
;=========================================================================
$arrayS1ListaPosicao[$eS1T10TelaCapaTerminalClickBotaoFecharX] 	= 757
$arrayS1ListaPosicao[$eS1T10TelaCapaTerminalClickBotaoFecharY] 	= 617
;========================================================================

;Selecionar Opcao Desconto De Mkt Lista suspensa
;=========================================================================
$arrayS1ListaPosicao[$eS1T10OpcaoDescontoMktClickX] 	= 492
$arrayS1ListaPosicao[$eS1T10OpcaoDescontoMktClickY] 	= 216
;========================================================================

;Selecionar Opcao Historico Lista Suspensa
;=========================================================================
$arrayS1ListaPosicao[$eS1T10OpcaoHistoricoClickX] 	= 502
$arrayS1ListaPosicao[$eS1T10OpcaoHistoricoClickY] 	= 387
;========================================================================

;Tela Mkt Click botao Fechar
;=========================================================================
$arrayS1ListaPosicao[$eS1T9MktBotaoFecharX] 	= 549
$arrayS1ListaPosicao[$eS1T9MktBotaoFecharY] 	= 352
;========================================================================

;Tela Historico Paginacao Lateral Esquerda
;=========================================================================
$arrayS1ListaPosicao[$eS1T10HistoricoPaginacaoLateralEsquerdaX] 	= 25
$arrayS1ListaPosicao[$eS1T10HistoricoPaginacaoLateralEsquerdaY] 	= 595
;========================================================================

;Tela Historico Paginacao Lateral Esquerda
;=========================================================================
$arrayS1ListaPosicao[$eS1T9SelecionarCapaTerminalX] 	= 783
$arrayS1ListaPosicao[$eS1T9SelecionarCapaTerminalY] 	= 641
;========================================================================

;Tela Historico Paginacao Lateral Esquerda
;=========================================================================
$arrayS1ListaPosicao[$eS1T10TelaTerminaisSelecionarX] 	= 161
$arrayS1ListaPosicao[$eS1T10TelaTerminaisSelecionarY] 	= 668
;=========================================================================

;Tela Selecionar Terminal Pesquisa
;=========================================================================
$arrayS1ListaPosicao[$eS1T7SelecionarTerminalX] 	= 91
$arrayS1ListaPosicao[$eS1T7SelecionarTerminalY] 	= 396
;=========================================================================

;Tela Click Botao Limpar
;=========================================================================
$arrayS1ListaPosicao[$eS1T7CapaTerminalBotaoLimparX] 	= 197
$arrayS1ListaPosicao[$eS1T7CapaTerminalBotaoLimparY] 	= 315
;=========================================================================

;Tela Lista Assinaturas
;=========================================================================
$arrayS1ListaPosicao[$eS1T10MensagemPesquisaNaoEncontradaInicioX] 	= 310
$arrayS1ListaPosicao[$eS1T10MensagemPesquisaNaoEncontradaInicioY]	= 277
$arrayS1ListaPosicao[$eS1T10MensagemPesquisaNaoEncontradaFimX]		= 598
$arrayS1ListaPosicao[$eS1T10MensagemPesquisaNaoEncontradaFimY]		= 318
;=========================================================================

;Tela Click Botao Limpar
;=========================================================================
$arrayS1ListaPosicao[$eS1T7CapaTerminalBotaoPesquisarX] 	= 97
$arrayS1ListaPosicao[$eS1T7CapaTerminalBotaoPesquisarY] 	= 315
;=========================================================================

;Tela Capa Terminal Fechar
;=========================================================================
$arrayS1ListaPosicao[$eS1T5TelaListaTerminaisBotaoFecharX]	= 724
$arrayS1ListaPosicao[$eS1T5TelaListaTerminaisBotaoFecharY] 	= 655
;=========================================================================

;Tela Capa Conta Fechar
;=========================================================================
$arrayS1ListaPosicao[$eS1T5TelaCapaContaBotaoFecharX]	= 713
$arrayS1ListaPosicao[$eS1T5TelaCapaContaBotaoFecharY] = 686
;=========================================================================

;Tela Assinaturas Botao Fechar
;=========================================================================
$arrayS1ListaPosicao[$eS1T5TelaAssinaturasBotaoFecharX]	= 709
$arrayS1ListaPosicao[$eS1T5TelaAssinaturasBotaoFecharY]	= 685
;=========================================================================

;Valida Preenchimento
;=========================================================================
$arrayS1ListaPosicao[$eS1T2ValidaRetornoPesquisaInicioX] 	= 27
$arrayS1ListaPosicao[$eS1T2ValidaRetornoPesquisaInicioY] 	= 352
$arrayS1ListaPosicao[$eS1T2ValidaRetornoPesquisaFimX]    	= 636
$arrayS1ListaPosicao[$eS1T2ValidaRetornoPesquisaFimY] 	= 371
;=========================================================================

;Tela Principal Botao Fechar
;=========================================================================
$arrayS1ListaPosicao[$eS1T1BotaoFecharTelaPrincipalX]	= 682
$arrayS1ListaPosicao[$eS1T1BotaoFecharTelaPrincipalY]	= 102
;=========================================================================

;Tela Historico Botao Fechar
;=========================================================================
$arrayS1ListaPosicao[$eS1T10ClickBotaoFecharX]	= 983
$arrayS1ListaPosicao[$eS1T10ClickBotaoFecharY]	= 484
;=========================================================================

;Valida Preenchimento
;=========================================================================
$arrayS1ListaPosicao[$eS1T1BotaoAssinaturaInicioX] 	= 350
$arrayS1ListaPosicao[$eS1T1BotaoAssinaturaInicioY] 	= 673
$arrayS1ListaPosicao[$eS1T1BotaoAssinaturaFimX]    	= 496
$arrayS1ListaPosicao[$eS1T1BotaoAssinaturaFimY] 	= 696
;=========================================================================

;Valida Botao Fechar Tela Terminais
;=========================================================================
$arrayS1ListaPosicao[$eS1T5TelaIdentBtnFecharInicioX] = 664
$arrayS1ListaPosicao[$eS1T5TelaIdentBtnFecharInicioY] = 638
$arrayS1ListaPosicao[$eS1T5TelaIdentBtnFecharFimX] = 790
$arrayS1ListaPosicao[$eS1T5TelaIdentBtnFecharFimY] = 673
;========================================================================


Func _MARCACAO_CORINGA($xPosicao,$yPosicao,$SleepTime)
   Sleep($SleepTime)
   MouseMove($xPosicao,$yPosicao)
   MouseClick("left")
EndFunc

Func _ShowStatus($Texto, $Titulo = "")
   IF $Texto <> "" THEN
	  ToolTip($Texto,(@DesktopWidth),(@DesktopHeight),$Titulo,1,4)
   Else
	  ToolTip("")
   EndIf
EndFunc
