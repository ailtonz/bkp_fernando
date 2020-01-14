#include <File.au3>
#include "bin\funcoes_UTIL_160830-1555.au3"
;~ #include "bin\RegrasAtlys.au3"
#include "bin\Capture2Text.au3"
#include "bin\PosVariaveis.au3"
;~ #include "OO_JSON.au3"
#include <Array.au3>
#include <Excel.au3>
#include <ScreenCapture.au3>
#include "ApiWebService.au3"

;~ WinActivate("Atlys Global Solution - \\Remote")
;~ _ScreenCapture_Capture(@MyDocumentsDir & "\PrintAtlys.jpg",33,445,1306,597)
;~ ShellExecute(@MyDocumentsDir & "\PrintAtlys.jpg")
;~ Exit

Global $tela = 0
Global $j = 0
Local $xP,$yP
Global $a
Global $sHttpServ = "http://portalcontestacao.com.br/"
Global $arrApiRota = "AzureOcr/api/service/image"
Global Const $hexaBarraGrid  = "D6D6CE"

#Region PreSets

	;Quando pressionar a tecla 'CTRL+END' sai da rotina
	HotKeySet("^{END}", "CaptureEND")
	HotKeySet("^{p}", "CaptureEND")

	; Liga a opção de eventos de controle
	Opt("GUIOnEventMode",1)

#EndRegion

#Region Main

;~ 	Local $buscarTela = _identificarTela($arrTela[$eAtlys],1000,5)
;~ 	WinActivate($buscarTela)
;~ 	Sleep(4000)

;~ MsgBox(0,"",$ResolucaoTela)

	Local $buscarTela = _identificarTela($arrTela[$eAtlys],1000,5)
	if $buscarTela = $arrTela[$eAtlys] Then
		Local $hWnd = WinWait($arrTela[$eAtlys], "", 10)
		Local $aClientSize = WinGetClientSize($hWnd)

		WinActivate($buscarTela)

		_ShowStatus("MAXIMIZAR - Atlys", $sApp[$eAutomatioName])
		_MARCACAO_CORINGA(741,12,$sApp[$eSleepTime])

;~ 		 ;Função responsavel por ler o arquivo MS EXCEL
	    _mainExcel()

;~ 		 ;Função responsavel por enviar e comando ao Atlys
		_PesquisaPagamentoPorNumeroDoLote()

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

Func _PesquisaPagamentoPorNumeroDoLote()

For $j = 0 To UBound($aArrayList) - 1

   If $aArrayList[$j][8] = "" Then

			   _ShowStatus("Pesquisa de pagamento", $sApp[$eAutomatioName])
			  _MARCACAO_CORINGA($sPesquisaPagamento[$ePesquisarInicioX],$sPesquisaPagamento[$ePesquisarInicioY],$sApp[$eSleepTime])
	  ;~ 		_MARCACAO_CORINGA(210,68,$sApp[$eSleepTime])

			  _ShowStatus("MAXIMIZAR - Pesquisa de pagamento", $sApp[$eAutomatioName])
			  _MARCACAO_CORINGA(605,100,$sApp[$eSleepTime])

			  _ShowStatus("Detalhe do lote", $sApp[$eAutomatioName])
			  SendWait("{DOWN}",$sApp[$eSleepTime])
			  SendWait("{SPACE}",$sApp[$eSleepTime])
	  ;~ 		_MARCACAO_CORINGA(281,205,$sApp[$eSleepTime])

			  _ShowStatus("numero do lote", $sApp[$eAutomatioName])
			  SendWait("{TAB}",$sApp[$eSleepTime],4)
	  ;~ 		SendWait($arrLote[$eLote],$sApp[$eSleepTime])
			  SendWait($aArrayList[$j][1],$sApp[$eSleepTime])

			  _ShowStatus("pagamentos não aplicados", $sApp[$eAutomatioName])
			  SendWait("{TAB}",$sApp[$eSleepTime],7)
			  SendWait("{SPACE}",$sApp[$eSleepTime])

			  _ShowStatus("Pesquisar", $sApp[$eAutomatioName])
			  SendWait("{TAB}",$sApp[$eSleepTime],5)
			  SendWait("{SPACE}",$sApp[$eSleepTime])

			   WinWaitActive($buscarTela)

			   Global $iColor  = PixelGetColor(1316,578)
			   Global $HexaColor = Hex($iColor, 6)

			   If $HexaColor <> $hexaBarraGrid Then

					 While $HexaColor <> $hexaBarraGrid

				  		_ScreenCapture_Capture(@MyDocumentsDir & "\PrintAtlys.jpg",33,445,1306,597)
						Global $file = FileOpen(@MyDocumentsDir & "\PrintAtlys.jpg",$FO_BINARY)
;~ 						ShellExecute(@MyDocumentsDir & "\PrintAtlys.jpg")
;~ 						Exit
						Global $Text = FileRead($file)
						Global $image64 = _Base64Encode($Text)

						;Função responsavel por procurar conta
						_Procura_Conta($aArrayList[$j][2],$j)

						 MouseClick("LEFT",1317,594)
						 Sleep(100)

					 WEnd
;~ 				  Do
;~ 						ConsoleWrite($image64)
;~ 						Exit

;~ 						;Loop para clicar 8 vezes para pular de pagina
;~ 						For $t = 1 to 8

;~

;~ 						Next

;~ 				  Until $HexaColor = $hexaBarraGrid

				 Else
;~ 					 _MARCACAO_CORINGA(147,453,$sApp[$eSleepTime])
;~ 					 SendWait("{SPACE}",$sApp[$eSleepTime])
					 _ScreenCapture_Capture(@MyDocumentsDir & "\PrintAtlys.jpg",33,445,1306,597,False)
					 Global $file = FileOpen(@MyDocumentsDir & "\PrintAtlys.jpg",$FO_BINARY)

;~ 					 ShellExecute(@MyDocumentsDir & "\PrintAtlys.jpg")
;~ 					 Exit
					 Global $Text = FileRead($file)
					 Global $image64 = _Base64Encode($Text)
;~ 					 ConsoleWrite($image64)
;~ 					 Exit
;~ 					 #include "ApiWebService.au3"

					 ;Função responsavel por procurar conta
					 _Procura_Conta($aArrayList[$j][2],$j)

			   EndIf

   EndIf

Next

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

Func _Main()

   Local $arrayPesquisar[1]
		 $arrayPesquisar[0] = "Pesquisar"

	Local $buscarTela = _identificarTela($arrTela[$eAtlys],1000,5)
	if $buscarTela = $arrTela[$eAtlys] Then

		WinActivate($buscarTela)

		Send("{F4}")

		Local $checkRota = False
		For $r = 1 To 3
			ConsoleWrite("_ValidaTextoTela")
			;Valida se a palavra chave foi encontrada na posição mapeada
				if _ValidaTextoTela($arrayS1ListaPosicao[$eS1T1PesquisarInicioX], _
									$arrayS1ListaPosicao[$eS1T1PesquisarInicioY], _
									$arrayS1ListaPosicao[$eS1T1PesquisarFimX]   , _
									$arrayS1ListaPosicao[$eS1T1PesquisarFimY]   , _
									$arrayPesquisar, 5,1000,1) Then

					$checkRota = True
					ExitLoop
				EndIf
			_reiniciarPesquisaAtlys()
		Next

		If $checkRota  = True Then

			;Criar Log Sucesso
			ConsoleWrite("----------- Identificou Sistema Iniciado" & @CRLF)

			Return True
		Else
			_KillAtlys()
		EndIf
	Endif

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

Func _mainExcel()

   ;INSTANCIAR OBJETO
   Global $oExcelObj = ObjCreate("Excel.Application")
   $oExcelObj.visible = True

   ;ABRIR WORKBOOK
   Global $oWbk =  $oExcelObj.WorkBooks.Open(@ScriptDir & "\Relatório Robo.xlsx")

   ;CARREGAR WORKSHEET CADASTRO - A planilha deve ser a primeira da esquerda
   Global $sht =  $oWbk.Worksheets(1)

   ;QUANTIDADE DE LINHAS
   Global $iRowCount = $sht.UsedRange.Rows.Count

   ;QUANTIDADE COLUNAS
   Global $iColumnCount = $sht.UsedRange.Columns.Count

   ;POSICIONAR APOS A ULTIMA COLUNA PARA STATUS
;~	Global $iLastColumn = _Excel_ColumnToLetter($iColumnCount+1)

   ;RANGE
   Global $iRange = "A2" & ":I" & $iRowCount
   ;Global $iRange = "A2" & ":BC" & $iRowCount

   ;ARRAY DO RANGE
   Global $aArrayList = _Excel_RangeRead($oWbk,$sht,$iRange,1,Default)

;~    _ArrayDisplay($aArrayList)

   ;_ArrayDisplay($aArrayList)

   ;Valida se a lista é um Array e se possuí dados
   If Not IsArray($aArrayList) Or Ubound($aArrayList) = 0 then
	  MsgBox($MB_ICONERROR,$sApp[$eAutomatioName],"Não foram identificados registros no arquivo selecionado!")
	  _Exit()
   EndIf

EndFunc

;Fecha a planilha e encerra a aplicação Excel
;============================================
Func _PLANILHA_FECHAR()
	$oWbk.close
	$oExcelObj.quit
	$oExcelObj = Null
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