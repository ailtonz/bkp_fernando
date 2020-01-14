#include <FileConstants.au3>

Local $sHttpServ = "http://172.16.0.11/"
Local $arrApiRota = "AzureOcr/api/service/image"
Global Const $hexaColorTelaAtlys  = "D6D6CE"

	Local $buscarTela = _identificarTela($arrTela[$eAtlys],1000,5)
   if $buscarTela = $arrTela[$eAtlys] Then
		Local $hWnd = WinWait($arrTela[$eAtlys], "", 10)
		Local $aClientSize = WinGetClientSize($hWnd)

	  If WinExists("Atlys Global Solution - \\Remote") Then

		 WinActivate("Atlys Global Solution - \\Remote")

		 Local $iColor  = PixelGetColor(1314,477)
		 Local $HexaColor = Hex($iColor, 6)

			;Condicao para verificar se a cor da barra é igual
			If $HexaColor <> $hexaColorTelaAtlys Then

			   Do
				  ;Loop para clicar 8 vezes para pular de pagina
				  For $t = 1 to 8

					MouseClick("LEFT",1314,489)
					Sleep(100)

				  Next

						$PrintGrid =  _ScreenCapture_Capture(@MyDocumentsDir & "\PrintAtlys.jpg",33,445,1306,597,False)
						Local $file = FileOpen(,$FO_BINARY)
						Local $Text = FileRead($file)
						Local $image64 = _Base64Encode($Text)

						#include "ApiWebService.au3"
						Local  $jsonToSendSTR = '{"Base64":"' & $image64 & '","Extension":".jpg"}'

						Local $Json = apiCall($sHttpServ,$arrApiRota,"POST",$jsonToSendSTR,"",$eTech,1,"","")

						Local $js = _OO_JSON_Init()
						Local $jsObj = $js.strToObject($Json)

				  Sleep(1000)

			   Until $HexaColor = $hexaColorTelaAtlys

			Else

						$PrintGrid =  _ScreenCapture_Capture(@MyDocumentsDir & "\PrintAtlys.jpg",33,445,1306,597,False)
						Local $file = FileOpen(,$FO_BINARY)
						Local $Text = FileRead($file)
						Local $image64 = _Base64Encode($Text)

						#include "ApiWebService.au3"
						Local  $jsonToSendSTR = '{"Base64":"' & $image64 & '","Extension":".jpg"}'

						Local $Json = apiCall($sHttpServ,$arrApiRota,"POST",$jsonToSendSTR,"",$eTech,1,"","")

						Local $js = _OO_JSON_Init()
						Local $jsObj = $js.strToObject($Json)

			EndIf

   EndIf




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
