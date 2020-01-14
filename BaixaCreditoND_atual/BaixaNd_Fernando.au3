#include <File.au3>
;~ #include "bin\funcoes_UTIL_160830-1555.au3"
;#include "bin\Capture2Text.au3"
#include "bin\PosVariaveis.au3"
#include "_getResultOcr.au3"
#include <Array.au3>
#include <Excel.au3>
#include <ScreenCapture.au3>
#include "ApiWebService.au3"
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <WindowsConstants.au3>
#include <GDIPlus.au3>
#include <Date.au3>


;-------------------------------------------
;Mapa de Colunas da Planilha Insumos do Robo
Global Enum 	$eDataLote 		_
			,	$eNrLote 		_
			,	$eNumeroConta 	_
			,	$eVlr 			_
			,	$eRef			_
			,	$eDtPgto		_
			,	$eCodBanco		_
			,	$eNsr			_
			,	$eLogProc1		_
			,	$eDtLog1		_
			,	$eLogProc2		_
			,	$eDtLog2

Global Const $poscursor  = 7
Global Const $hexaBarraGrid  = "D6D6CE"
Global Const $fileImg = @ScriptDir & "\prints\PrintAtlys.jpg"
Global Const $fileImg2 = @ScriptDir & "\prints\PrintAtlysConverted.jpg"

Global Const $fileImgCheck = @ScriptDir & "\prints\PrintAtlys"
Global Const $fileImgCheck2 = @ScriptDir & "\prints\PrintAtlysConverted"

Global $fileImgCoringa = @ScriptDir & "\prints\PrintAtlys_"
Global $fileImgCoringa2 = @ScriptDir & "\prints\PrintAtlysConverted_"

Global Enum $eContaRetorno, $eDtPgtoRetorno, $eVlrRetorno

Global $ApplicationId = "TechForOps_TLF_BPO"
Global $Password = "HyxdUzcN4oMVZSOsdCEcQKlK"

$toEncode = $ApplicationId & ":" & $Password
$token = _Base64Encode($toEncode)

Global $tela = 0
Global $j = 0
Local $xP,$yP
Global $a
Global $sHttpServ = "http://portalcontestacao.com.br/"
Global $arrApiRota = "AzureOcr/api/service/image"

;~ Global $FilePath = @ScriptDir & "\prints\PrintAtlys.jpg"


#Region PreSets
   ;Quando pressionar a tecla 'CTRL+END' sai da rotina
   HotKeySet("^{END}", "CaptureEND")
#EndRegion


;~ _TestArea()
_LimparPastaPrint()
_IniciarProcesso()




Func _TestArea()

	FileDelete(@ScriptDir & "\prints\*.jpg")
Sleep(100000000)
;~ 	Global $txtNumeroConta = "0310591130"
;~ 	Global $ContaTransferencia = GUICtrlRead($txtNumeroConta)
;~     Global $i
;~ 	Local $caminho = "C:\Users\f.salvador.de.sa\Desktop\Workspace\Baixa Credito\BaixaCreditoND_atual\BaixaCreditoND_atual\Teste Robo_Lista3.xlsx"
;~ 	_mainExcel($caminho)

;~ 	_ArrayDisplay($aArrayValues);

;~ 	For $i = 0 To Ubound($aArrayValues) - 1
;~ 		$oSht.Cells($i + 2,8).value = "OK"
;~ 		$oSht.Cells($i + 2,9).value = _Now()
;~ 		MsgBox(0,"Teste","Teste Stop",1)
;~ 	Next

;~ 	For $i = 0 To Ubound($aArrayValues) - 1
;~ 		$oSht.Cells($i + 2,10).value = "OK"
;~ 		$oSht.Cells($i + 2,11).value = _Now()
;~ 		MsgBox(0,"Teste","Teste Stop",1)
;~ 	Next
;~ 	Sleep(100000000)


EndFunc

Func _IniciarProcesso()

   _exibeInterface()


EndFunc

Func _exibeInterface()

   ;Liga a opção de eventos de controle
   Opt("GUIOnEventMode",1)

   ;Criação do Form
   ;============================================================
   Global $frmBaixaNd = GUICreate("Baixa Nota de Débito", 378, 200, -1, -1)
   GUISetOnEvent($GUI_EVENT_CLOSE, "_CloseForm") ; botão X do Form

   ;Caixas de Texto
   ;=====================================================
   Global $lblUser = GUICtrlCreateLabel("Numero Conta Baixa", 8, 30, 96, 17, 0)
   Global $txtNumeroConta = GUICtrlCreateInput("0310591130", 6, 44, 161, 21) ;Txt Usuário Com Dígito

   ;Criação Agrupador Seleção Planilha
   ;========================================================
   $grpPlanProc = GUICtrlCreateGroup("Selecionar Planilha", 0, 80, 377, 81)
   GUICtrlCreateGroup("", -99, -99, 1, 1)

   ;Caixa de Texto
   ;======================================================
   Local $caminho = ""
   Global $txtFileName = GUICtrlCreateInput($caminho, 6, 100, 367, 21)


   ;Botão Selecionar
   ;======================================================
   Global $btnSelecionar = GUICtrlCreateButton("Selecionar", 5, 130, 75, 25)
   GUICtrlSetOnEvent($btnSelecionar, "carregarPlanilha")

   ;Criação dos Botões do Form
   ;========================================================
   Global $btnOk = GUICtrlCreateButton("Start", 6, 170, 75, 25)
   GUICtrlSetOnEvent($btnOk, "_Main")

   Global $btnCancel = GUICtrlCreateButton("Cancel", 87, 170, 75, 25)
   GUICtrlSetOnEvent($btnCancel, "_CloseForm")

   GUISetState(@SW_SHOW)

   ;-- Aguarda ação do usuário
   While 1
	  Sleep(10); comando para nao usar muito a CPU
   WEnd
EndFunc

Func _Main() ;Inicio



   Global $ContaTransferencia = GUICtrlRead($txtNumeroConta)
   Global $i
   _mainExcel(GUICtrlRead($txtFileName))


	;_ArrayDisplay($aArrayValues)

	For $i = 0 To Ubound($aArrayValues) - 1

		if $aArrayValues[$i][7] = "" Then
			Local $arrExec[Ubound($aArrayValues,2)]
			  For $e = 0 To Ubound($aArrayValues,2) -1
						$arrExec[$e] = StringStripWS($aArrayValues[$i][$e],$STR_STRIPLEADING + $STR_STRIPTRAILING + $STR_STRIPSPACES)
			  Next
			ConsoleWrite("Entrando_PesquisaPagamentoPorNumeroDoLote" & @CRLF)
			Local $result
			$result = _PesquisaPagamentoPorNumeroDoLote($arrExec,$ContaTransferencia)

			$oSht.Cells($i + 2,8).value = $result
			$oSht.Cells($i + 2,9).value = _Now()

			ConsoleWrite("Saindo_PesquisaPagamentoPorNumeroDoLote" & @CRLF)

		EndIf

	Next


	For $i = 0 To Ubound($aArrayValues) - 1
		if $aArrayValues[$i][9] = "" Then

			Local $arrExec[Ubound($aArrayValues,2)]
			For $e = 0 To Ubound($aArrayValues,2) -1
				$arrExec[$e] = StringStripWS($aArrayValues[$i][$e],$STR_STRIPLEADING + $STR_STRIPTRAILING + $STR_STRIPSPACES) ;transpoe a tabela excel, cada linha em uma coluna
			Next

	;~ 		_ArrayDisplay($arrExec)
			ConsoleWrite("Entrando_PesquisaReversaoDePagamento" & @CRLF)
			Local $result
			$result = _PesquisaReversaoDePagamento($arrExec,$ContaTransferencia)

			$oSht.Cells($i + 2,10).value = $result
			$oSht.Cells($i + 2,11).value = _Now()

			ConsoleWrite("Saindo_PesquisaReversaoDePagamento" & @CRLF)
		EndIf
	Next

EndFunc

#Region Main

	Local $buscarTela = _identificarTela($arrTela[$eAtlys],1000,5)
	if $buscarTela = $arrTela[$eAtlys] Then
		Local $hWnd = WinWait($arrTela[$eAtlys], "", 10)
		Local $aClientSize = WinGetClientSize($hWnd)

		WinActivate($buscarTela)

		_ShowStatus("MAXIMIZAR - Atlys", $sApp[$eAutomatioName])
		_MARCACAO_CORINGA(741,12,$sApp[$eSleepTime])

;~ 		 ;Função responsavel por ler o arquivo MS EXCEL
;~ 	    _mainExcel()

;~ 		 ;Função responsavel por enviar e comando ao Atlys
;~ 		_PesquisaPagamentoPorNumeroDoLote()

;~ 		_RevisaoDePagamento()

;~ 		 _ReversaoPagmento()


	EndIf

;~ 	_MARCACAO_CORINGA(247,345,$sApp[$eSleepTime])
;~ 	SendWait($sApp[$uName],$sApp[$eSleepTime]*2)

#EndRegion

Func _RevisaoDePagamento()

		_ShowStatus("Reversão de pagamento", $sApp[$eAutomatioName])
		_MARCACAO_CORINGA(275,61,$sApp[$eSleepTime])

		_ShowStatus("Maximizar", $sApp[$eAutomatioName])
		_MARCACAO_CORINGA(536,104,$sApp[$eSleepTime])

		_ShowStatus("pesquisar", $sApp[$eAutomatioName])
		SendWait("0222046478",$sApp[$eSleepTime])
		SendWait("{TAB}",$sApp[$eSleepTime])
		SendWait("{SPACE}",$sApp[$eSleepTime])
		Sleep(5000)

		Do
			$MotivoReversao = _Teste(224,343,309,354,1000)
			If $MotivoReversao <> "" Then
			   SendWait("{TAB}",$sApp[$eSleepTime])
			   _MARCACAO_CORINGA(947,347,$sApp[$eSleepTime])
			   _ReversaoPagmento()
			EndIf

	    Until $MotivoReversao <> ""

EndFunc

Func _ReversaoPagmento()

;Global $MesAnoInformado = InputBox("INFORMAÇÃO PARA PERCORRER A LISTA","Informe o mês e ano para percorra o robô execute até a data informada")

Global $PosYInicial = 340
Global $PosYFinal = 356

   		 Do
			;$Reversao = _Teste2(866,$PosYInicial,1019,$PosYFinal,1000,$MesAnoInformado)
			$Data = _ValidaData(63,339,104,354,100)
			If $Reversao = "" And $Data =  Then
			   _PAGAMENTOERRO($Reversao)
			EndIf
		 Until $Reversao = "" Or $Data = $MesAnoInformado

EndFunc

Func _RetornaTextoOcr($tipoDado,$fileImage)

   Local $txtRetorno
   Local $erro
   Local $aRetornoTela1[8][3]

   ;Tipo Dado 1 = Grid Busca de Pagamento
   $postData = FileRead($fileImage)
;~    MsgBox(0,"",$postData)

   if  _GetResultOcr($token, $postData, "PortugueseBrazilian", "txt", $txtRetorno, $erro) = false then
	  MsgBox(0,"Erro Conversao","Erro de conversão de imgagem" & @CRLF & $erro)
	  Exit
   Else
	  ConsoleWrite($txtRetorno & @CRLF)
	  if $tipoDado = 1 Then

		 Local $procura = StringSplit(_ALLTRIM($txtRetorno),@CRLF)
		 Local $iRow = 0
		 For $i = 1 to $procura[0]
			if StringLen($procura[$i]) > 0 Then

			   Local $registro = _filtroCaracteres($procura[$i])
			   ConsoleWrite($registro & @CRLF)

			EndIf
		 Next

		 Return $registro
	  EndIf
   Endif

EndFunc

Func _PesquisaReversaoDePagamento($aRegistro,$NumeroContaTransferencia)

Local $data
Local $valor
Local $vlrexcel
Local $day
Local $month
Local $year
Local $excelday
Local $excelmonth
Local $excelyear
Local $index = 1
Local $bEncontrou = False
Local $stepscursor
Local $baixado = False
Local $buscarTela = _identificarTela($arrTela[$eAtlys],1000,5)
Local $status

   if $buscarTela = $arrTela[$eAtlys] Then
	  WinActivate($buscarTela)
   Else
	  MsgBox(0,"Atlys Nâo Encontrado!","Sistema atlys não identificado!" & @crlf & "Favor iniciar o sistema e iniciar o robo novamente!")
	  _PLANILHA_FECHAR()
	  Exit
   EndIf

      ;if $aRegistro[$eLogProc1]
   ;Verifica se o Status do Processo foi executado anteriormente
   ;-------------------------------------------------------------
   if $aRegistro[$eLogProc2] = "" Then

		;Selecionar a oção Pesquisa de Pagamento
	  _ShowStatus("Reversão de pagamento do cliente", $sApp[$eAutomatioName])
	  _MARCACAO_CORINGA($sReversaoPagamento[$eOpcaoPesquisaPgtoX],$sReversaoPagamento[$eOpcaoPesquisaPgtoY],$sApp[$eSleepTime]) ;Seleciona o icone de reversão de pgto.

	  _ShowStatus("MAXIMIZAR - Reversão de pagamento", $sApp[$eAutomatioName])
	  _MARCACAO_CORINGA($sReversaoPagamento[$eMaximizarTelaPgtoX],$sReversaoPagamento[$eMaximizarTelaPgtoY],$sApp[$eSleepTime])

	  _ShowStatus("preenchimento numero conta", $sApp[$eAutomatioName])
	  SendWait($NumeroContaTransferencia,$sApp[$eSleepTime])

	  _ShowStatus("Selecionar Pesquisar", $sApp[$eAutomatioName])
	  SendWait("{ENTER}",3000)

		;Tem que esperar carregar a lista após a pesquisa

		sleep(1000)

		if WinExists("Solução Global Atlys  - \\Remote") Then
			SendWait("{ENTER}",1000)
			MsgBox(0,"Conta Inexistente","Conta não localizada")
		Else
			;Rotina para validar o carregamento da pesquisa
			;~ 					 ;Verificar o pixel preto
			Local $iColorPx  = PixelGetColor(1250, 346)

			For $z = 0 TO 1000
				MouseClick("left",1250, 346)
				sleep(1000)
				$iColorPx  = PixelGetColor(1250, 346)
				if $iColorPx = 0x000000 Then
					;MouseClick("left",1250, 346)
					SendWait("{DOWN}",500,1)
					SendWait("{DOWN}",500,1)
					SendWait("{DOWN}",500,1)
					SendWait("{DOWN}",500,1)
					SendWait("{DOWN}",500,1)
					SendWait("{DOWN}",500,1)
					SendWait("{DOWN}",500,1)
					ExitLoop
				EndIf
			Next

		EndIf


		Local $listaImagem

	For $z = 0 TO 1000
		Local $imagem =  $fileImgCoringa & $aRegistro[$eNumeroConta]& "_" & $z &  ".jpg"
		Local $imagem2 =  $fileImgCoringa2 & $aRegistro[$eNumeroConta]& "_" & $z &  ".jpg"

		 if FileExists($imagem) Then
			FileDelete($imagem)
		 EndIf

		 if FileExists($imagem2) Then
			FileDelete($imagem2)
		 EndIf

		_ScreenCapture_Capture($imagem, _
											 $sPrintReversao[$eRegContasX], _
											 $sPrintReversao[$eRegContasY], _
											 $sPrintReversao[$eRegContasFimX]	 , _
											 $sPrintReversao[$eRegContasFimY],False)

		 _GDIPlus_Startup()
			SaveBmp2JPG($imagem,$imagem2)
		 _GDIPlus_Shutdown()


		ConsoleWrite("$imagem2 = " & $imagem2 & @CRLF)
		 $listaImagem = _ChamaOcr2(1,$imagem2) ;retorna uma lista com as contas que forem mostradas na tabela de pesquisa de contas.
		 ConsoleWrite("Conta = " & $aRegistro[$eNumeroConta] & @CRLF)



		For $y = 0 To 1000

				$data =  $listaImagem[0][0]

				$dataexcel = StringLeft($aRegistro[$eDtPgto],4) & "-" & StringLeft(StringRight($aRegistro[$eDtPgto],4),2) & "-" & StringRight($aRegistro[$eDtPgto],2)
				$dataocr = StringLeft($data,4)& "-" & StringLeft(StringRight($data,4),2) & "-" & StringRight($data,2)

				$iDays = _DateDiff("D", $dataexcel, $dataocr)

			For $o = 0 TO Ubound($listaImagem)-1 ;Procura a conta na listagem aberta, após a pesquisa
				$vlrexcel = StringReplace(StringReplace($aRegistro[$eVlr],".",""),",","")
				$data =  $listaImagem[$o][0]
				$status = $listaImagem[$o][2]
				$valor = _RTRIM(_LTRIM(StringReplace(StringReplace($listaImagem[$o][1],".",""),",","")))

				ConsoleWrite("Valor = " & $valor & @CRLF)
				ConsoleWrite("ValorExcel = " & $vlrexcel & @CRLF)
				ConsoleWrite("Data OCR = " & $data & @CRLF)

				ConsoleWrite("Instr Pagamento = " & StringInStr($status,"PAGAMENTO") & @CRLF)
				if _validaPercentualAcerto($data, $aRegistro[5],80) = True and StringInStr($status,"PAGAMENTO") = 0 Then

				   if _validaPercentualAcerto($valor,$vlrexcel,80) = True Then
					  $stepscursor = $poscursor - $index
					  $bEncontrou = True
					  $baixado = True
					  _NavegaLista($stepscursor)
					  ExitLoop
				   EndIf
				EndIf
				$index = $index + 1
			Next

			if $bEncontrou = false Then
				SendWait("{DOWN}",100,1)
				SendWait("{DOWN}",100,1)
				SendWait("{DOWN}",100,1)
				SendWait("{DOWN}",100,1)
				SendWait("{DOWN}",100,1)
				SendWait("{DOWN}",100,1)
				SendWait("{DOWN}",100,1)
				$index = 1
				;MsgBox(0,"Mensagem","Lista movida, pois não encontrou o valor",3)
				ExitLoop
			EndIf

			if $baixado = True then
				ExitLoop
			EndIf
;~ 			MsgBox(0,"Loop interno","Valor $iDays = " & $iDays,3)
			if $iDays <= -2 Then
				MsgBox(0,"Saindo","Saindo do loop",3)
				ExitLoop
			EndIf
		Next

			if $baixado = True then
				ExitLoop
			EndIf

			if $iDays <= -2 Then
				ExitLoop
			EndIf
	Next

	MouseClick("left",1353, 95);ok

   EndIf

   		if $bEncontrou = True Then
			Return "Ok"
		Else
			Return "Não encontrado"
		EndIf

EndFunc

Func _NavegaLista($stepscursor)
	;MsgBox(0,"Navega lista","Inicio Navega lista",3)
	For $i = 0 to $stepscursor
		SendWait("{UP}",500,1)
	Next

		;Clicar em Reverter
		MouseClick("left",57, 529)
		sleep(500)
		SendWait("{TAB}",500,1)
		SendWait("PAGAMENTO",500,1)
		SendWait("{TAB}",500,1)
		MouseClick("left",154, 634);ok
		;SendWait("{ENTER}",500,1)
		sleep(5000)
		MouseClick("left",305, 378);ok
		;MouseClick("left",666, 378);cancelar

		;MsgBox(0,"Voltando","Voltando Cursor",5)
	For $i = 0 to $stepscursor
		SendWait("{DOWN}",500,1)
	Next

	Return True

EndFunc
;#######################################################################
Func _PesquisaPagamentoPorNumeroDoLote($aRegistro,$NumeroContaTransferencia)

   Local $buscarTela = _identificarTela($arrTela[$eAtlys],1000,5)
   Local $bEncontrou = False

   if $buscarTela = $arrTela[$eAtlys] Then
	  WinActivate($buscarTela)
   Else
	  MsgBox(0,"Atlys Nâo Encontrado!","Sistema atlys não identificado!" & @crlf & "Favor iniciar o sistema e iniciar o robo novamente!")
	  _PLANILHA_FECHAR()
	  Exit
   EndIf

   ;-------------------------------------------------------------
   ;if $aRegistro[$eLogProc1]
   ;Verifica se o Status do Processo foi executado anteriormente
   ;-------------------------------------------------------------
   if $aRegistro[$eLogProc1] = "" Then

	  ;Selecionar a oção Pesquisa de Pagamento
	  _ShowStatus("Pesquisa de pagamento", $sApp[$eAutomatioName])
	  _MARCACAO_CORINGA($sPesquisaPagamento[$eOpcaoPesquisaPgtoX],$sPesquisaPagamento[$eOpcaoPesquisaPgtoY],$sApp[$eSleepTime])

	  _ShowStatus("MAXIMIZAR - Pesquisa de pagamento", $sApp[$eAutomatioName])
	  _MARCACAO_CORINGA($sPesquisaPagamento[$eMaximizarTelaPgtoX],$sPesquisaPagamento[$eMaximizarTelaPgtoY],$sApp[$eSleepTime])

	  _ShowStatus("Detalhe do lote", $sApp[$eAutomatioName])
	  SendWait("{DOWN}",$sApp[$eSleepTime])
	  SendWait("{SPACE}",$sApp[$eSleepTime])

	  _ShowStatus("pagamentos não aplicados", $sApp[$eAutomatioName])
	  SendWait("{TAB}",$sApp[$eSleepTime])
	  SendWait("{SPACE}",$sApp[$eSleepTime])

;~ 	  _ShowStatus("preenchimento data lote", $sApp[$eAutomatioName])
	  SendWait("{TAB}",$sApp[$eSleepTime],2)
;~ 	  SendWait(_trataDataRelatorio($aRegistro[$eDataLote]),$sApp[$eSleepTime])

	  _ShowStatus("preenchimento numero lote", $sApp[$eAutomatioName])
	  SendWait("{TAB}",$sApp[$eSleepTime])
	  SendWait($aRegistro[$eNrLote],$sApp[$eSleepTime])

	  _ShowStatus("Selecionar Pesquisar", $sApp[$eAutomatioName])
	  SendWait("{TAB}",$sApp[$eSleepTime],2)
	  SendWait("{ENTER}",3000)


	  Do

		 if FileExists($fileImg) Then
			FileDelete($fileImg)
		 EndIf

		 if FileExists($fileImg2) Then
			FileDelete($fileImg2)
		 EndIf

		 ;Verificar o pixel preto
		 Local $iColorPx  = PixelGetColor(1270, 453)
		 if $iColorPx = 0x000000 Then
					MouseClick("left",1263, 452)
					Sleep(500)
				send("{Space}")
		 EndIf

		 Sleep(2000)
		ConsoleWrite("Tirando print" & @CRLF)

		Local $imagem =  $fileImgCoringa & $aRegistro[$eNumeroConta]& ".jpg"
		Local $imagem2 =  $fileImgCoringa2 & $aRegistro[$eNumeroConta]& ".jpg"

		 _ScreenCapture_Capture($imagem, _
											 $sPesquisaPagamento[$eMapaPrintTLoteInicioX], _
											 $sPesquisaPagamento[$eMapaPrintTLoteInicioY], _
											 $sPesquisaPagamento[$eMapaPrintTLoteFimX]	 , _
											 $sPesquisaPagamento[$eMapaPrintTLoteFimY],False)

		 _GDIPlus_Startup()

		 SaveBmp2JPG($imagem,$imagem2)

		 _GDIPlus_Shutdown()

		 ConsoleWrite("Salvou print em = " & $imagem2 & @CRLF)

		 MouseClick("left",200, 450) ;clica na primeira linha da lista de contas

		 ConsoleWrite("Iniciando OCR" & @CRLF)
 		 Local $listaImagem = _ChamaOcr(1,$imagem2) ;retorna uma lista com as contas que forem mostradas na tabela de pesquisa de contas.
			;_ArrayDisplay($listaImagem)
		 Local $index = 0
		 Local $bEncontrou = False

		 For $z = 0 TO Ubound($listaImagem)-1 ;Procura a conta na listagem aberta, após a pesquisa

			Local $Conta = Number($aRegistro[$eNumeroConta])
			ConsoleWrite("Conta no Excel = " & $Conta & @CRLF)
			ConsoleWrite("Conta no OCR = " & $listaImagem[$z][$eContaRetorno] & @CRLF)
			ConsoleWrite("Len Conta no Excel = " & StringLen($aRegistro[$eNumeroConta]) & @CRLF)
			ConsoleWrite("Len Conta no OCR = " & StringLen($listaImagem[$z][$eContaRetorno]) & @CRLF)
;~ 			MsgBox(0,"Compara Excel x OCR = ",$Conta & " - " & $listaImagem[$z][$eContaRetorno],2)

			if _validaPercentualAcerto($listaImagem[$z][$eContaRetorno], $Conta) = True Then
			ConsoleWrite("_validaPercentualAcerto Conta Passou" & @CRLF)
			Local $vlr = StringReplace(StringReplace($aRegistro[$eVlr],".",""),",","")
			local $chars = StringLen($vlr)

			$listaImagem[$z][$eVlrRetorno] = StringRight($listaImagem[$z][$eVlrRetorno],$chars)

			   if _validaPercentualAcerto($listaImagem[$z][$eVlrRetorno],$vlr,50) = True Then

				  $bEncontrou = True
				  ExitLoop
			   EndIf
			EndIf
			$index = $index + 1
		 Next

		 if $bEncontrou = False Then
			 ConsoleWrite("Conta não encontrada, iniciando loop")
			;Loop para clicar 8 vezes para pular de pagina
;~ 			For $t = 1 to 8
				_MARCACAO_CORINGA(1315,584,100)
;~ 				_MARCACAO_CORINGA($sPesquisaPagamento[$eClickPaginarTLoteX],$sPesquisaPagamento[$eClickPaginarTLoteY],100)
;~ 			Next

			Local $iColor  = PixelGetColor($sPesquisaPagamento[$eColorGridTLoteX],$sPesquisaPagamento[$eColorGridTLoteY])
			Local $HexaColor = Hex($iColor, 6)

		 Else
				MouseClick("left",200, 450) ;clica na primeira linha da lista de contas
				ConsoleWrite("Apertando DOWN por " & $index & " vezes")
			   SendWait("{DOWN}",$sApp[$eSleepTime],$index) ;sendwait(tecla, tempo de sleep, quantas vezes repetir o comando)

 			   _ShowStatus("Selecionando Registro Na Grid", $sApp[$eAutomatioName])

				MouseClick("left",315, 645)
;~ 			   _ShowStatus("Selecionar Visualizar detalhes do pagamento", $sApp[$eAutomatioName])
;~ 			   SendWait("{ENTER}",$sApp[$eSleepTime])

			   _ShowStatus("Selecionar Botão de Confirmação", $sApp[$eAutomatioName])
			   SendWait("{TAB}",$sApp[$eSleepTime])

			   _ShowStatus("Confirmar Seleção", $sApp[$eAutomatioName])
			   SendWait("{ENTER}",$sApp[$eSleepTime])

			   _ShowStatus("Informar Numero Conta de Transferencia", $sApp[$eAutomatioName])
			   SendWait($NumeroContaTransferencia,$sApp[$eSleepTime])

			   _ShowStatus("Selecionar Ok", $sApp[$eAutomatioName])
			   SendWait("{TAB}",$sApp[$eSleepTime],2)

			   _ShowStatus("Selecionar Ok", $sApp[$eAutomatioName])
			   SendWait("{ENTER}",$sApp[$eSleepTime])

			   _ShowStatus("Confirmar", $sApp[$eAutomatioName])
			   SendWait("{ENTER}",$sApp[$eSleepTime])

			   _ShowStatus("Selecionar Atribuicao Padrao", $sApp[$eAutomatioName])
			   SendWait("{Down}",$sApp[$eSleepTime])

			   _ShowStatus("Confirmar Ok", $sApp[$eAutomatioName])
			   SendWait("{TAB}",$sApp[$eSleepTime],4)

			   _ShowStatus("Confirmar Ok", $sApp[$eAutomatioName])
			   SendWait("{ENTER}",$sApp[$eSleepTime])

			   _ShowStatus("Confirmar Ok", $sApp[$eAutomatioName])
			   SendWait("{ENTER}",$sApp[$eSleepTime])

			;MsgBox(0,"Escrevendo Excel","Escreveu ok na linha" & $i + 1 & " Coluna " & $eLogProc1)


			ExitLoop
		 EndIf

	  Until $HexaColor = $hexaBarraGrid


	  _MARCACAO_CORINGA($sPesquisaPagamento[$eFecharTelaPgtoX],$sPesquisaPagamento[$eFecharTelaPgtoY],100)
   EndIf

		if $bEncontrou = True Then
			Return "Ok"
		Else
			Return "Não encontrado"
		EndIf

    ConsoleWrite("Fim _PesquisaPagamentoPorNumeroDoLote")

EndFunc

Func _Procura_Conta($ContaPlanilha,$LinhaPlanilha)

Global $jsonToSendSTR = '{"Base64":"' & $image64 & '","Extension":".jpg"}'

Global $Json = apiCall($sHttpServ,$arrApiRota,"POST",$jsonToSendSTR,"",$eTech,1,"","")

Global $js = _OO_JSON_Init()
Global $jsObj = $js.strToObject($Json)

 For $i = 0 to $jsObj.regions.length() -1
   For $a = 0 TO $jsObj.regions.item($i).lines.length()-1

;~ 	  MsgBox(0,"",StringMid($aArrayList[$j][2],1,5))

	  If StringLen($jsObj.regions.item($i).lines.item($a).words.item(0).text) = 10 And _
		 StringIsInt($jsObj.regions.item($i).lines.item($a).words.item(0).text) Then

	   $ContaJson = StringReplace(StringReplace($jsObj.regions.item($i).lines.item($a).words.item(0).text,"B","8"),"å·³","8")

		 If $ContaPlanilha = $ContaJson Then

			MsgBox(0,"", $ContaJson)

			Sleep(2000)
		   _ShowStatus("visualizar/alterar", $sApp[$eAutomatioName])
		   SendWait("{TAB}",2000,2)

			$ColorGrid = PixelGetColor(54,449)

			If $ColorGrid = "0xFFFFFF" Then
			   SendWait("{SPACE}",$sApp[$eSleepTime])
			   SendWait("{DOWN}",$sApp[$eSleepTime],$a)
			   SendWait("^c",1000)
			Else
;~ 			   SendWait("{SPACE}",$sApp[$eSleepTime])
			   SendWait("{TAB}",$sApp[$eSleepTime])
			   SendWait("^c",$sApp[$eSleepTime])
			EndIf

			_ShowStatus("visualizar/alterar pagamento não aplicado", $sApp[$eAutomatioName])
			SendWait("0222046478",$sApp[$eSleepTime])
			SendWait("{TAB}",$sApp[$eSleepTime],2)
			Sleep(2000)

		   SendWait("{ENTER}",$sApp[$eSleepTime])
		   WinWaitActive("Solução Global Atlys  - \\Remote")
		   SendWait("{ENTER}",$sApp[$eSleepTime])

		   _ShowStatus("Efetuar atribuições", $sApp[$eAutomatioName])
		   WinWaitActive("Efetuar atribuições - \\Remote")
		   SendWait("{DOWN}",$sApp[$eSleepTime],1)
		   SendWait("{TAB}",$sApp[$eSleepTime],4)
		   SendWait("{ENTER}",$sApp[$eSleepTime])

		   _ShowStatus("Efetuar atribuições", $sApp[$eAutomatioName])
		   WinWaitActive("Atribuições reais - \\Remote")
		   SendWait("{TAB}",$sApp[$eSleepTime],2)
		   SendWait("{ENTER}",$sApp[$eSleepTime])

		   _Excel_RangeWrite($oWbk, $oWbk.Worksheets(1),"OK","I" & $LinhaPlanilha + 2)

			    Local $TelaPesquisa
			    While $TelaPesquisa <> "Conta"
				  $TelaPesquisa = _Teste(54,418,96,434,1000)
				  If $TelaPesquisa = "Conta" Then
					  _MARCACAO_CORINGA(1351,97,$sApp[$eSleepTime])
					  Sleep(4000)
					  $TelaPesquisa = ""
					  ExitLoop
				  EndIf
			   WEnd
			EndIf

	  ElseIf StringLen($jsObj.regions.item($i).lines.item($a).words.item(0).text) > 10 Then

		  _Excel_RangeWrite($oWbk, $oWbk.Worksheets(1),"CONTA NÃO VALIDA, POR FAVOR EXECUTAR O PROCESSO MANUALMENTE","I" & $LinhaPlanilha + 2)

	  EndIf
   Next
Next

EndFunc

Func _PAGAMENTOERRO($Reversao)

   If $Reversao = "" Then
	  SendWait("{TAB}",$sApp[$eSleepTime],1)
	  SendWait("{ENTER}",$sApp[$eSleepTime],1)
   EndIf

	_ShowStatus("Reverter Pagamento", $sApp[$eAutomatioName])
	SendWait("{TAB}",$sApp[$eSleepTime],1)
    ;SendWait("{ENTER}",$sApp[$eSleepTime])

     _ShowStatus("Motivo reversão", $sApp[$eAutomatioName])
	 _MARCACAO_CORINGA(997,529,$sApp[$eSleepTime])
     SendWait("{DOWN}",$sApp[$eSleepTime],4)
	 SendWait("{ENTER}",$sApp[$eSleepTime])
     SendWait("{TAB}",$sApp[$eSleepTime],3)
	 SendWait("{ENTER}",$sApp[$eSleepTime])
     WinWaitActive("Solução Global Atlys  - \\Remote")
     SendWait("{TAB}",$sApp[$eSleepTime])
	 SendWait("{ENTER}",$sApp[$eSleepTime])
	 _Excel_RangeWrite($oWbk, $oWbk.Worksheets(1),"Pagamento Revertido","J" & $j + 2)
     _MARCACAO_CORINGA(1352,96,$sApp[$eSleepTime])
	 _ReversaoPagmento()

EndFunc

Func _identificarTela($nomeTela,$milSecWait,$qtdBusca = 50,$iBuscaExata = 0)
   ;====================================================================================
   ;Funcao retorna o nome da tela pesquisada dentro da lista de telas abertas do sistema
   ;====================================================================================
   For $i = 1 To $qtdBusca

	  Local $aList = WinList()

	  if $iBuscaExata = 0 Then
		 Local $iIndex = _ArraySearch($aList,$nomeTela,Default,Default,Default)
	  Else
		 Local $iIndex = _ArraySearch($aList,$nomeTela,Default,Default,Default,$iBuscaExata)
	  EndIf

	  if $iIndex > -1 Then
		 Return $aList[$iIndex][0]
	  EndIf

	  Sleep($milSecWait)
   Next

   Return -1
EndFunc

Func _reiniciarPesquisaAtlys()

   Local $arrayPesquisar[1]
		 $arrayPesquisar[0] = "Pesquisar"

;~    Local $buscarTela = _identificarTela($arrTela[$eAtlys],1000)
;~    if $buscarTela = $arrTela[$eAtlys] Then
	  WinActivate($arrTela[$eAtlys])

	  MouseClick("LEFT",$arrayS1ListaPosicao[$eS1T3TelaNotasBotaoFecharX],$arrayS1ListaPosicao[$eS1T3TelaNotasBotaoFecharY],1)
	  Sleep(500)

	  MouseClick("LEFT",$arrayS1ListaPosicao[$eS1T6PgtoAjusteMinimizarTelaX],$arrayS1ListaPosicao[$eS1T6PgtoAjusteMinimizarTelaY],1)
	  Sleep(500)

	  MouseClick("LEFT",$arrayS1ListaPosicao[$eS1T1BotaoFecharTelaPrincipalX],$arrayS1ListaPosicao[$eS1T1BotaoFecharTelaPrincipalY],1)
	  Sleep(500)

	  MouseClick("LEFT",$arrayS1ListaPosicao[$eS1T2TelaResumoCobrancaFecharX],$arrayS1ListaPosicao[$eS1T2TelaResumoCobrancaFecharY],1)
	  Sleep(500)

	  Send("{F4}")

	  if _ValidaTextoTela($arrayS1ListaPosicao[$eS1T1PesquisarInicioX]	, _
							 $arrayS1ListaPosicao[$eS1T1PesquisarInicioY]	, _
							 $arrayS1ListaPosicao[$eS1T1PesquisarFimX]    , _
							 $arrayS1ListaPosicao[$eS1T1PesquisarFimY]   	, _
							 $arrayPesquisar, 5,1000) Then
		 Return True
	  EndIf
;~    EndIf
   Return False
EndFunc

Func _KillAtlys()

   if WinExists($arrTela[$eAtlysSolution]) Then
	  WinActivate($arrTela[$eAtlysSolution])
	  Sleep(100)

	  WinKill($arrTela[$eAtlysSolution])
	  Send("{ENTER}")
   EndIf
EndFunc

Func _mainExcel($sFilename,$sSheetName = null, $indexSheet = 1)
;Função que faz a chamada do arquivo Excel selecionado pelo usuário e carrega as informações a serem processadas

   Local $iFileExists  = FileExists($sFilename)

   If Not $iFileExists Then
	  MsgBox($MB_ICONERROR, $iMsgTitle, "Arquivo " & $sFilename & " não encontrado.")
	  Exit
   EndIf

   ;INSTANCIAR OBJETO
   Global $appExcel = ObjCreate("Excel.Application")
   $appExcel.visible = true	;mostra o excel.

   ;ABRIR WORKBOOK
   Global $oWbk =  $appExcel.WorkBooks.Open($sFilename)

   ;CARREGAR WORKSHEET
   ;Local $aWorkSheet = _Excel_SheetList($oDocument)
   if $sSheetName = Null Then
	  $sSheetName = $indexSheet
   EndIf

   Global $oSht =  $oWbk.Worksheets($sSheetName)

   ;QUANTIDADE DE LINHAS
   Global $iRowCount = _qtdLinExcel($oSht,"A")

   ;Range de Pesquisa
   Global $sRange = "A2:L" & $iRowCount

   ; ARRAY DO RANGE
   Global $aArrayValues = _Excel_RangeRead($oWbk,$oSht,$sRange,1,Default)

   ;_ArrayDisplay($aArrayValues)

   If Ubound($aArrayValues,0) = 0 then
	  MsgBox($MB_ICONWARNING,"","Não Foram Encontrados Valores Na Planilha Selecionada!")
	  Exit
   Endif

   Return $aArrayValues

EndFunc

Func _PLANILHA_FECHAR()
	$oWbk.Save
	$oWbk.close
	$appExcel.quit
	$appExcel = Null
 EndFunc

 Func _Base64Encode($Data, $LineBreak = 76)


	Local $Opcode = "0x5589E5FF7514535657E8410000004142434445464748494A4B4C4D4E4F505152535455565758595A6162636465666768696A6B6C6D6E6F707172737475767778797A303132333435363738392B2F005A8B5D088B7D108B4D0CE98F0000000FB633C1EE0201D68A06880731C083F901760C0FB6430125F0000000C1E8040FB63383E603C1E60409C601D68A0688470183F90176210FB6430225C0000000C1E8060FB6730183E60FC1E60209C601D68A06884702EB04C647023D83F90276100FB6730283E63F01D68A06884703EB04C647033D8D5B038D7F0483E903836DFC04750C8B45148945FC66B80D0A66AB85C90F8F69FFFFFFC607005F5E5BC9C21000"

	Local $CodeBuffer = DllStructCreate("byte[" & BinaryLen($Opcode) & "]")
	DllStructSetData($CodeBuffer, 1, $Opcode)

	$Data = Binary($Data)
	Local $Input = DllStructCreate("byte[" & BinaryLen($Data) & "]")
	DllStructSetData($Input, 1, $Data)

	$LineBreak = Floor($LineBreak / 4) * 4
	Local $OputputSize = Ceiling(BinaryLen($Data) * 4 / 3)
	$OputputSize = $OputputSize + Ceiling($OputputSize / $LineBreak) * 2 + 4

	Local $Ouput = DllStructCreate("char[" & $OputputSize & "]")
    Local $sDLL = DllOpen("user32.dll")
    if Not $sDLL = 0 Then
	  DllCall($sDLL, "none", "CallWindowProc", "ptr", DllStructGetPtr($CodeBuffer), _
													"ptr", DllStructGetPtr($Input), _
													"int", BinaryLen($Data), _
													"ptr", DllStructGetPtr($Ouput), _
													"uint", $LineBreak)
	  DllClose($sDLL)
	  Return DllStructGetData($Ouput, 1)
   Else
	  Return False
   EndIf
EndFunc

Func _CloseForm()
   Exit
EndFunc

Func CaptureEND()
   Switch @HotKeyPressed ; The last hotkey pressed.
	  Case "^{end}" ; String is the {end} hotkey.
		 Local $iRetorno = MsgBox(4,"Pause","Deseja parar o script?")
		 If $iRetorno = 6 Then
			Exit
		 EndIf
   EndSwitch
EndFunc

Func carregarPlanilha()
   local $tmpArqTexto = _SelecionarArquivo("*.xls*")
   GUICtrlSetData($txtFileName,$tmpArqTexto)
EndFunc

Func _SelecionarArquivo($sExtensao)
    ; Create a constant variable in Local scope of the message to display in FileOpenDialog.
    Local Const $sMessage = "Selecione um único arquivo."

    ; Display an open dialog to select a file.
    Local $sFileOpenDialog = FileOpenDialog($sMessage, @ScriptDir & "\", "All (" & $sExtensao & ")", $FD_FILEMUSTEXIST)
    If @error Then
        ; Display the error message.
        MsgBox($MB_SYSTEMMODAL, "", "Nenhum arquivo foi selecionado.")

        ; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
        FileChangeDir(@ScriptDir)
    Else
        ; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
        FileChangeDir(@ScriptDir)

        ; Replace instances of "|" with @CRLF in the string returned by FileOpenDialog.
        $sFileOpenDialog = StringReplace($sFileOpenDialog, "|", @CRLF)

        ; Display the selected file.
;~         MsgBox($MB_SYSTEMMODAL, "", "Você escolheu o seguinte arquivo:" & @CRLF & $sFileOpenDialog)

		Return $sFileOpenDialog
    EndIf
 EndFunc

Func _qtdLinExcel($sht,$strCol)
  Return ($sht.Range($strCol & 65536).End(-4162).Row)
EndFunc

Func _ChamaOcr2($tipoDado,$fileImage)

   Local $txtRetorno
   Local $erro
   Local $aRetornoTela1[1][3]

   Local $dd
   Local $mm
   Local $yyyy
   Local $yyyymmdd

   Local $procuras
   Local $stringlen
   Local $tratado
   Local $procuradinheiro
   Local $valor
   Local $status



   ;Tipo Dado 1 = Grid Busca de Pagamento
   $postData = FileRead(FileOpen($fileImage,16))
   ;$postData = FileRead($fileImage)

	;MsgBox(0,"",StringLen($postData))

   if  _GetResultOcr($token, $postData, "PortugueseBrazilian", "txt", $txtRetorno, $erro) = false then
	  MsgBox(0,"Erro Conversao","Erro de conversão de imgagem" & @CRLF & $erro)
	  Exit
   Else
	  ConsoleWrite($txtRetorno & @CRLF) ;txtretorno retorna todo o texto encontrado pela OCR
	  if $tipoDado = 1 Then

		 Local $procura = StringSplit(_ALLTRIM($txtRetorno),@CRLF)

		 if UBound($procura)>0 Then

			 Local $validtens[1]

			 ;_ArrayDisplay($procura)

			 local $j = 0
			 Local $ilinha
			 Local $aRet


			For $i = 0 to UBound($procura) - 1
				ConsoleWrite("Valor $procura[" & $i & "] = " & $procura[$i] & @CRLF )
				$ilinha = StringInStr($procura[$i],"PAGAMENTO COM ERRO")

					local $ibarra = StringInStr($procura[$i],"/")
					;$status = StringLeft(StringInStr($tratado,"PAGAMENTO COM ERRO")

					If $ibarra > 5 Then ;Limpar caracteres estranhos
						$procura[$i] = StringRight($procura[$i],StringLen($procura[$i])-3)
					EndIf

					if $procura[$i] <> "" and StringLen($procura[$i]) > 20 and StringInStr($procura[$i],"$") > 1 Then
						$validtens[$j] = $procura[$i]
						ReDim $validtens[UBound($validtens) + 1]
						$j=$j+1
					endIf
			Next
				;_ArrayDisplay($validtens)
			For $i = 0 to UBound($validtens) - 1
				;## Refatorar este código
				$procuras = StringInStr($validtens[$i],"$")
				$stringlen = StringLen($validtens[$i])
				$tratado = StringRight($validtens[$i],$stringlen-$procuras)
				$procuradinheiro = StringInStr($tratado,"DINHEIRO") - 2
				$valor = StringLeft($tratado,$procuradinheiro - 1)
				$valor = StringLeft(StringRight($validtens[$i],$stringlen-$procuras),$procuradinheiro)

				$status = StringLeft(StringRight($validtens[$i],StringInStr($validtens[$i],"PAGAMENTO COM ERRO")),19)

				;MsgBox(0,"Teste",$status)

				$dd = StringLeft($validtens[$i],2)
				$mm = StringRight(StringLeft($validtens[$i],5),2)
				$yyyy = StringRight(StringLeft($validtens[$i],10),4)
				$yyyymmdd =  $yyyy & $mm & $dd



				If $valor <> "" Then
				$aRetornoTela1[$i][0] = $yyyymmdd
				$aRetornoTela1[$i][1] = $valor
				$aRetornoTela1[$i][2] = $status

				ReDim $aRetornoTela1[UBound($aRetornoTela1) + 1][3]

				EndIf

			Next




				;_ArrayDisplay($aRetornoTela1)
				ConsoleWrite("Fim OCR" & @CRLF )
				Return $aRetornoTela1
		 EndIf
		 Return False
	  EndIf
   Endif

EndFunc

Func _LimparPastaPrint()
	FileDelete(@ScriptDir & "\prints\*.jpg")
EndFunc

Func _GetDataAndValor($linha)

Local $values[2]
local $dd = StringLeft($linha,2)
local $mm = StringRight(StringLeft($linha,5),2)
local $yyyy = StringRight(StringLeft($linha,10),4)
local $yyyymmdd =  $yyyy & $mm & $dd


local $procuras = StringInStr($linha,"$")
local $stringlen = StringLen($linha)
local $tratado = StringRight($linha,$stringlen-$procuras)
local $procuradinheiro = StringInStr($tratado,"DINHEIRO")
local $valor = StringLeft($tratado,$procuradinheiro - 1)

$values[0] = $yyyymmdd
$values[1] = $valor

;_ArrayDisplay($values)

return $values

EndFunc

Func _ChamaOcr($tipoDado,$fileImage)

   Local $txtRetorno
   Local $erro
   Local $aRetornoTela1[15][3]

    $postData = FileRead(FileOpen($fileImage,16))

   if  _GetResultOcr($token, $postData, "PortugueseBrazilian", "txt", $txtRetorno, $erro) = false then
	  MsgBox(0,"Erro Conversao","Erro de conversão de imgagem" & @CRLF & $erro)
	  Exit
   Else
	  ConsoleWrite($txtRetorno & @CRLF)
	  if $tipoDado = 1 Then

		 Local $procura = StringSplit(_ALLTRIM($txtRetorno),@CRLF)
;~ 		 MsgBox(0,"TxtRetorno","Array $txtRetorno",1)
		 ;_ArrayDisplay($txtRetorno)
		 Local $iRow = 0
		 For $i = 1 to $procura[0]
			if StringLen($procura[$i]) > 0 Then

			   Local $registro = _filtroCaracteres($procura[$i])

			   ConsoleWrite("$registro = " & $registro & @CRLF)
			   ConsoleWrite("$iRow = " & $iRow & @CRLF)

			   $aRetornoTela1[$iRow][$eContaRetorno] 	= StringLeft($registro,10)
			   $aRetornoTela1[$iRow][$eDtPgtoRetorno] 	= StringMid($registro,11,8)
			   $aRetornoTela1[$iRow][$eVlrRetorno] 		= StringMid($registro,19,StringLen($registro)-18)
			   $iRow = $iRow + 1
			EndIf
		 Next

		 Return $aRetornoTela1
	  EndIf
   Endif

EndFunc

Func _trataDataRelatorio($data)
   ;Data do Relatório vem como yyyymmdd - No sistema é necessário o preenchimento ddmmyyyy
   Return StringRight($data,2) & StringMid($data,5,2) & StringLeft($data,4)
EndFunc

Func SendWait($text,$sleep,$repeat=1,$format="%s",$flag=0)
   for $j = 1 to $repeat
	 Send(StringFormat($format,$text),$flag)
	 Sleep($sleep/$repeat)
  Next
EndFunc

func _ALLTRIM($sString, $sTrimChars='')

	;  Trim from left first, then right

	$sTrimChars = StringReplace( $sTrimChars, "%%whs%%", " " & chr(9) & chr(11) & chr(12) & @CRLF )
	local $sStringWork = ""

	$sStringWork = _LTRIM($sString, $sTrimChars)
	if $sStringWork <> "" then
		$sStringWork = _RTRIM($sStringWork, $sTrimChars)
	endif
	return $sStringWork

endfunc

func _LTRIM($sString, $sTrimChars=' ')

	$sTrimChars = StringReplace( $sTrimChars, "%%whs%%", " " & chr(9) & chr(11) & chr(12) & @CRLF )
	local $nCount, $nFoundChar
	local $aStringArray = StringSplit($sString, "")
	local $aCharsArray = StringSplit($sTrimChars, "")

	for $nCount = 1 to $aStringArray[0]
		$nFoundChar = 0
		for $i = 1 to $aCharsArray[0]
			if $aCharsArray[$i] = $aStringArray[$nCount] then
				$nFoundChar = 1
			EndIf
		next
		if $nFoundChar = 0 then return StringTrimLeft( $sString, ($nCount-1) )
	next
endfunc

func _RTRIM($sString, $sTrimChars=' ')

	$sTrimChars = StringReplace( $sTrimChars, "%%whs%%", " " & chr(9) & chr(11) & chr(12) & @CRLF )
	local $nCount, $nFoundChar
	local $aStringArray = StringSplit($sString, "")
	local $aCharsArray = StringSplit($sTrimChars, "")

	for $nCount = $aStringArray[0] to 1 step -1
		$nFoundChar = 0
		for $i = 1 to $aCharsArray[0]
			if $aCharsArray[$i] = $aStringArray[$nCount] then
				$nFoundChar = 1
			EndIf
		next
		if $nFoundChar = 0 then return StringTrimRight( $sString, ($aStringArray[0] - $nCount) )
	next
 endfunc

 Func _filtroCaracteres($caract)

   local $filtro = "1234567890"
   local $result = ""

   For $i = 1 To StringLen($caract)
	  $p = StringInStr($filtro, StringMid($caract, $i, 1))
	  If $p > 0 Then
		 $result = $result & StringMid($caract, $i, 1)
	  EndIf
   Next

   Return $result
EndFunc

Func _gerarInsumo()




Local $aRetornoTela1[8][3]

	  $aRetornoTela1[0][0]  = "6400063167"
	  $aRetornoTela1[0][1]	= "07/03/2018"
	  $aRetornoTela1[0][2]  = "2998"

	  $aRetornoTela1[1][0]  = "6400068492"
	  $aRetornoTela1[1][1]	= "07/03/2018"
	  $aRetornoTela1[1][2]  = "2720"

	  $aRetornoTela1[2][0]  = "6400068512"
	  $aRetornoTela1[2][1]	= "07/03/2018"
	  $aRetornoTela1[2][2]	= "466278"

	  $aRetornoTela1[3][0]  = "6400064182"
	  $aRetornoTela1[3][1]	= "07/03/2018"
	  $aRetornoTela1[3][2]	= "18015"

	  $aRetornoTela1[4][0]  = "6400068507"
	  $aRetornoTela1[4][1]	= "07/03/2018"
	  $aRetornoTela1[4][2]	= "494336"

	  $aRetornoTela1[5][0]  = "6400068502"
	  $aRetornoTela1[5][1]	= "07/03/2018"
	  $aRetornoTela1[5][2]	= "532905"

	  $aRetornoTela1[6][0]  = "6400068482"
	  $aRetornoTela1[6][1]	= "07/03/2018"
	  $aRetornoTela1[6][2]	= "2923"

	  $aRetornoTela1[7][0]  = "6400063822"
	  $aRetornoTela1[7][1]	= "07/03/2018"
	  $aRetornoTela1[7][2]	= "567247"

	  Return $aRetornoTela1
   EndFunc

Func SaveBmp2JPG($Bitmap, $sSave, $iQuality = 80.9) ;coded by UEZ 2013
    If Not IsPtr($Bitmap) Then

        $Bitmap = _GDIPlus_ImageLoadFromFile($Bitmap)
        If @error Then Return SetError(1, 0, 0)
	EndIf

    Local $sCLSID = _GDIPlus_EncodersGetCLSID("JPG")
    Local $tParams = _GDIPlus_ParamInit(1)
    Local $tData = DllStructCreate("int Quality")
    Local $pData = DllStructGetPtr($tData)
    Local $pParams = DllStructGetPtr($tParams)

	;$iW = width
	;$iH = height
	Local $iW = _GDIPlus_ImageGetWidth($Bitmap)
	Local $iH = _GDIPlus_ImageGetHeight($Bitmap)

	Local $hBitmap_Scaled = _GDIPlus_ImageResize($Bitmap, $iW + 100 , $iH + 100)
	;_WinAPI_DeleteObject($Bitmap)


    DllStructSetData($tData, "Quality", $iQuality)
    _GDIPlus_ParamAdd($tParams, $GDIP_EPGQUALITY, 1, $GDIP_EPTLONG, $pData)
    ;If Not _GDIPlus_ImageSaveToFileEx($Bitmap, $sSave, $sCLSID, $pParams) Then Return SetError(2, 0, 0)
	If Not _GDIPlus_ImageSaveToFileEx($hBitmap_Scaled, $sSave, $sCLSID, $pParams) Then Return SetError(2, 0, 0)
    Return True
 EndFunc

Func _validaPercentualAcerto($string1, $string2,$percentualOk = 95)

   if StringLen($string1) = StringLen($string2) Then
	  $qtdAcertos = 0
	  For $i = 1 TO StringLen($string1)
		 if StringMid($string1,$i,1) = StringMid($string2,$i,1) Then
			$qtdAcertos = $qtdAcertos + 1
		 EndIf
	  Next

	  ;MsgBox(0,"",$string1 & @CRLF & $string2 & @CRLF & $qtdAcertos/StringLen($string1) * 100)
	  if $qtdAcertos/StringLen($string1) * 100 >= $percentualOk Then
		 Return True
	  EndIf
   EndIf

   Return False

EndFunc

;~ Func _PopUpErrorCheck()

;~ 		 _ScreenCapture_Capture($fileImgCheck, _
;~ 											 320, _
;~ 											 277, _
;~ 											 662	 , _
;~ 											 339,False)

;~ 		 _GDIPlus_Startup()
;~ 		 SaveBmp2JPG($fileImgCheck,$fileImgCheck2)
;~ 		 _GDIPlus_Shutdown()


;~ 		 Local $listaImagem = RetornaTextoOcr(1,$fileImgCheck2) ;_ChamaOcr(1,$fileImgCheck2) ;retorna uma lista com as contas que forem mostradas na tabela de pesquisa de contas.

;~ 		 _ArrayDisplay($listaImagem)

;~ 		 ;Local Teste = RetornaTextoOcr(1,$fileImgCheck2)

;~ 		 ;MsgBox(0,"Teste",Teste)

;~ EndFunc



;~ 		For $row = 1 to UBound($tabExcel)
;~
;~

;~ 		  If $tabExcel[$row] = True And $PosArray[$index] = True Then
;~ 		     $PosArray[$index] = $oSht.Range("E" & $row).Value And ;Comparando valor
;~ 			 $PosArray[$index] = $oSht.Range("A" & $row).Value Then ;Comparando data

;~ 				$oSht.Range("K" & $row).Value = "OK"
;~ 				$oSht.Range("L" & $row).Value = @MDAY & "/" & @MON & "/" & @YEAR

;~ 		  EndIf

;~ 		Next


;~ 		MsgBox(0,"",$testesheet)


	  ;_ArrayDisplay($aArrayValues)



	  ;~ Local $listaImagem
;~ $listaImagem = _ChamaOcr2(1,$fileImgCheck2) ;retorna uma lista com as contas que forem mostradas na tabela de pesquisa de contas.

;~ Local $container


;~ For $i = 0 to UBound($listaImagem)-1 ;Loop que varre a lista do OCR
;~ 	$container = _GetDataAndValor($listaImagem[$i])
;~ 	_ArrayDisplay($container)
;~ Next


;~ Sleep(1000000000)






;~ 			ConsoleWrite("Validou, Conta no OCR = " & $listaImagem[$z][$eVlrRetorno] & @CRLF)
;~ 			ConsoleWrite("Validou, Conta no Excel = " & StringReplace(StringReplace($aRegistro[$eVlr],".",""),",","") & @CRLF)
