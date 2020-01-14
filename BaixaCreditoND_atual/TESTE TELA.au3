#include <ScreenCapture.au3>
#include <File.au3>
#include <ScreenCapture.au3>
#include "bin\funcoes_UTIL_160830-1555.au3"
#include "bin\RegrasAtlys.au3"
#include "bin\Capture2Text.au3"
#include "bin\PosVariaveis.au3"

Example()

Func Example()
;~     Local $hGUI

;~ Local $buscarTela = _identificarTela($arrTela[$eAtlys],1000,5)
;~ 	if $buscarTela = $arrTela[$eAtlys] Then
;~ 		Local $hWnd = WinWait($arrTela[$eAtlys], "", 10)
;~ 		Local $aClientSize = WinGetClientSize($hWnd)

;~ 	   ; Create GUI
;~ 	   $hWnd = GUICreate("Screen Capture",500,500,True)

;~ 	   ; Capture window
;~ 	   _ScreenCapture_CaptureWnd(@MyDocumentsDir & "\GDIPlus_Image.jpg", $hGUI)

;~ 	  ShellExecute(@MyDocumentsDir & "\GDIPlus_Image.jpg")
;~    EndIf

WinActivate("Atlys Global Solution - \\Remote")
Sleep(1000)

_ScreenCapture_Capture("C:\Users\wesley.valadao\Documents\BaixaCreditoND\teste.jpg",32,442,1308,601,False)
ShellExecute("C:\Users\wesley.valadao\Documents\BaixaCreditoND\teste.jpg")

EndFunc   ;==>Example
