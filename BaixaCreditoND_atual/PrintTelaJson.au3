#include <FileConstants.au3>
#include "ApiWebService.au3"

Local $sHttpServ = "http://172.16.0.11/"
Local $arrApiRota = "AzureOcr/api/service/image"

Local $buscarTela = _identificarTela($arrTela[$eAtlys],1000,5)
	if $buscarTela = $arrTela[$eAtlys] Then
		Local $hWnd = WinWait($arrTela[$eAtlys], "", 10)
		Local $aClientSize = WinGetClientSize($hWnd)

		WinActivate($buscarTela)

		_ShowStatus("MAXIMIZAR - Atlys", $sApp[$eAutomatioName])
		_MARCACAO_CORINGA(741,12,$sApp[$eSleepTime])

		;Função responsavel por ler o arquivo MS EXCEL
	    _mainExcel()


;~ 		 ;Função responsavel por enviar e comando ao Atlys
;~ 		_PesquisaPagamentoPorNumeroDoLote()

		 Local $iColor  = PixelGetColor(1314,477)
		 Local $HexaColor = Hex($iColor, 6)

		 If $HexaColor <> $hexaColorTelaAtlys Then

			Do
			   ;Loop para clicar 8 vezes para pular de pagina
			   For $t = 1 to 8

				 MouseClick("LEFT",1314,489)
				 Sleep(100)

			   Next

				  ;------------------------------------------------------------------------------------------------------
				 ;CHAMAR API DO JSON PAR A TIRAR PRINT DA TELA
				  $PrintGrid =  _ScreenCapture_Capture(@MyDocumentsDir & "\PrintAtlys.jpg",33,445,1306,597,False)
				  Local $file = FileOpen($PrintGrid ,$FO_BINARY)
				  Local $Text = FileRead($file)
				  Local $image64 = _Base64Encode($Text)

				  #include "ApiWebService.au3"
				  Local  $jsonToSendSTR = '{"Base64":"' & $image64 & '","Extension":".jpg"}'

				  Local $Json = apiCall($sHttpServ,$arrApiRota,"POST",$jsonToSendSTR,"",$eTech,1,"","")

				  Local $js = _OO_JSON_Init()
				  Local $jsObj = $js.strToObject($Json)

			Until $HexaColor = $hexaColorTelaAtlys

		  Else

			  ;------------------------------------------------------------------------------------------------------
				 ;CHAMAR API DO JSON PAR A TIRAR PRINT DA TELA
				  $PrintGrid =  _ScreenCapture_Capture(@MyDocumentsDir & "\PrintAtlys.jpg",33,445,1306,597,False)
				  Local $file = FileOpen($PrintGrid ,$FO_BINARY)
				  Local $Text = FileRead($file)
				  Local $image64 = _Base64Encode($Text)

				  #include "ApiWebService.au3"
				  Local  $jsonToSendSTR = '{"Base64":"' & $image64 & '","Extension":".jpg"}'

				  Local $Json = apiCall($sHttpServ,$arrApiRota,"POST",$jsonToSendSTR,"",$eTech,1,"","")

				  Local $js = _OO_JSON_Init()
				  Local $jsObj = $js.strToObject($Json)

				 ;-------------------------------------------------------------------------------------------------------

		  EndIf

	EndIf

Func _PesquisaPagamentoPorNumeroDoLote()

For $j = 0 To UBound($aArrayList) - 1
   For $i = 0 to $jsObj.regions.length() -1
	  For $a = 0 TO $jsObj.regions.item($i).lines.length()-1

			If $aArrayList[$j][2] = $jsObj.regions.item($i).lines.item($a).words.item(0).text AND StringStripWS($aArrayList[$j][8],$STR_STRIPALL) = ""  Then

			  ;MsgBox(0,"",$jsObj.regions.item($i).lines.item($a).words.item(0).text)

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

			   Sleep(2000)
			   _ShowStatus("visualizar/alterar", $sApp[$eAutomatioName])
			   SendWait("{TAB}",2000,2)

			   $ColorGrid = PixelGetColor(54,449)

			   If $ColorGrid = "0xFFFFFF" Then
				  SendWait("{SPACE}",$sApp[$eSleepTime])
				  SendWait("{DOWN}",$sApp[$eSleepTime],$a)
				  SendWait("^c",1000)
			   Else
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

			   _Excel_RangeWrite($oWbk, $oWbk.Worksheets(1),"OK","I" & $j + 2)

			   Do
				  $TelaPesquisa = _Teste(54,418,96,434,1000)
			   Until $TelaPesquisa = "Conta"

			EndIf
		 Next
	  Next
   Next

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
    if $sDLL = 1 Then
	  DllCall($sDLL, "none", "CallWindowProc", "ptr", DllStructGetPtr($CodeBuffer), _
													"ptr", DllStructGetPtr($Input), _
													"int", BinaryLen($Data), _
													"ptr", DllStructGetPtr($Ouput), _
													"uint", $LineBreak)
	  Return DllStructGetData($Ouput, 1)
   Else
	  Return False
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