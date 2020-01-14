#include <___OO_JSON.au3>
#include <ApiWebService.au3>
;~ #include <Base64.au3>
#include <OO_JSON.au3>
#include <WinHttp.au3>
#include <WinHttpConstants.au3>
#include <ScreenCapture.au3>
#include <PrintTelaJson.au3>
#include <FileConstants.au3>
#include "ApiWebService.au3"

;Quando pressionar a tecla 'CTRL+END' sai da rotina
HotKeySet("{END}", "CaptureEND")

; Liga a opção de eventos de controle
Opt("GUIOnEventMode",1)

Global Const $hexaColorTelaAtlys  = "D6D6CE"

If WinExists("Atlys Global Solution - \\Remote") Then

   WinActivate("Atlys Global Solution - \\Remote")

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
		    _ScreenCapture_Capture(@ScriptDir & "\PRINT PASSO 2.jpg",38,339,1301,492,False)
			$file = (@ScriptDir & "\PRINT PASSO 2.jpg")
			Local $Text = FileRead($file)
			Local $image64 = _Base64Encode($Text)

			Local  $jsonToSendSTR = '{"Base64":"' & $image64 & '","Extension":".jpg"}'

			Local $Json = apiCall($sHttpServ,$arrApiRota,"POST",$jsonToSendSTR,"",$eTech,1,"","")

			For $i = 0 to $jsObj.regions.length() -1
				For $a = 0 TO $jsObj.regions.item($i).lines.length()-1
					ConsoleWrite($jsObj.regions.item($i).lines.item($a).words.item(0).text & @CRLF)
				Next
			Next
			;~ 		    ShellExecute(@MyDocumentsDir & "\PrintAtlys.jpg")
		   ;-------------------------------------------------------------------------------------------------------

		 Sleep(1000)

	  Until $HexaColor = $hexaColorTelaAtlys

   	Else

		;------------------------------------------------------------------------------------------------------
		   ;CHAMAR API DO JSON PAR A TIRAR PRINT DA TELA
		   _ScreenCapture_Capture(@ScriptDir & "\PRINT PASSO 2.jpg",38,339,1301,492,False)
		    $PrintGrid =  (@ScriptDir & "\PRINT PASSO 2.jpg")
			Local $Text = FileRead( $PrintGrid)
			Local $image64 = _Base64Encode($Text)

			Local  $jsonToSendSTR = '{"Base64":"' & $image64 & '","Extension":".jpg"}'

			Local $Json = apiCall($sHttpServ,$arrApiRota,"POST",$jsonToSendSTR,"",$eTech,1,"","")

			For $i = 0 to $jsObj.regions.length() -1
				For $a = 0 TO $jsObj.regions.item($i).lines.length()-1
					ConsoleWrite($jsObj.regions.item($i).lines.item($a).words.item(0).text & @CRLF)
				Next
			Next
			;~ 		    ShellExecute(@MyDocumentsDir & "\PrintAtlys.jpg")
		   ;-------------------------------------------------------------------------------------------------------

	EndIf

EndIf

Func CaptureEND()
   Exit
EndFunc