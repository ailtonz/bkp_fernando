#include "bin\funcoes_UTIL_160830-1555.au3"

Global $arrTela[8]
Global Enum $eWinLogon,$eAtlys,$eNgin,$eVivoNet,$eAtlysSolution,$eBloqueioAtlys,$eCSO,$eAtlysErro

	$arrTela[$eAtlys] 			= "Reversão de pagamento do cliente"
	$arrTela[$eAtlysSolution] 	= "Logon do Atlys Global Solution - \\Remote"
	$arrTela[$eBloqueioAtlys] 	= "Bloqueio da Solução Global Atlys - \\Remote"
	$arrTela[$eAtlysErro]		= "Atlys ERROR"



Local $xP,$yP

;~ "Document1 - Word"

#Region PreSets

	;Quando pressionar a tecla 'CTRL+END' sai da rotina
	HotKeySet("^{END}", "CaptureEND")
	HotKeySet("^{p}", "CaptureEND")

	; Liga a opção de eventos de controle
	Opt("GUIOnEventMode",1)

#EndRegion


#Region Main


	_Main()


#EndRegion




Func _Main()


Local $buscarTela = _identificarTela($arrTela[$eAtlys],1000,5)

if $buscarTela = $arrTela[$eAtlys] Then

	Local $hWnd = WinWait($arrTela[$eAtlys], "", 10)
	Local $aClientSize = WinGetClientSize($hWnd)
	While 1

		local $aPos = MouseGetPos()
		ToolTip($aClientSize[0] & "x" & $aClientSize[1] & @CR & "x: " & $aPos[0] & @CR & "y: " & $aPos[1], Default, Default, Default, Default, 4)
		$xP = $aPos[0]
		$yP = $aPos[1]
		Sleep(100)

	WEnd
	EndIf

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