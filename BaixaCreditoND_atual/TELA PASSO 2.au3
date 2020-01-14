#include <ScreenCapture.au3>

If WinExists("Atlys Global Solution - \\Remote") Then

WinActivate("Atlys Global Solution - \\Remote")

;_ScreenCapture_Capture(38,339,1301,492, @ScriptDir & "\PRINT PASSO 2.jpg",False)

_ScreenCapture_Capture(@ScriptDir & "\PRINT PASSO 2.jpg",38,339,1301,492,False)
ShellExecute(@ScriptDir & "\PRINT PASSO 2.jpg")

EndIf