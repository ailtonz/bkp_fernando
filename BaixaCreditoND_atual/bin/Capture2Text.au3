Global $CaminhoOCR  = @ScriptDir & "\bin\Config\Capture2Text_v4.4.0_64bit\"
;~ Global $CaminhoOCR2 = @ScriptDir & "\Config\Capture2Text_v4.4.0_64bit\"

func _ValidaTextoTela($PosXInicial,$PosYInicial,$PosXFinal,$PosYFinal,$arrayValidar,$NumTentativas,$TempoEsperaTentativa,$iDebug = 0,$ExecMaximizar = 0)

   Run($CaminhoOCR & "Capture2Text.exe")

   Sleep(2000)

   For $i = 1 to $NumTentativas

	  Send("{ESC}")

	  MouseMove($PosXInicial,$PosYInicial,1)
	  Sleep(100)

	  ClipPut("")

	  Send("!q")
	  Sleep(100)

	  MouseMove($PosXFinal,$PosYFinal,1)
	  Sleep(100)

	  Send("!q")
	  Sleep(1000)

	  Global $sCaptura = ClipGet()

	  Send("{ESC}")

	  For $a = 0 TO Ubound($arrayValidar)-1
		 if StringInStr($sCaptura,$arrayValidar[$a],1) Then
			if $iDebug = 1 Then
			   MsgBox(0,"Debug Valida",$sCaptura)
			EndIf
			ProcessClose("Capture2Text.exe")
			Return True
		 EndIf
	  Next

	  if $ExecMaximizar =  1 Then
		Send("!{SPACE}x")
	  EndIf
	  Sleep($TempoEsperaTentativa)
   Next

   if $iDebug = 1 Then
	  MsgBox(0,"Debug Valida",$sCaptura)
   EndIf
   ProcessClose("Capture2Text.exe")
   Return False
EndFunc

func _ValidaPreenchimento($PosXInicial,$PosYInicial,$PosXFinal,$PosYFinal,$NumTentativas,$TempoEsperaTentativa,$iDebug = 0)

   Run($CaminhoOCR & "Capture2Text.exe")

   Sleep(2000)

   For $i = 1 to $NumTentativas

	  Send("{ESC}")

	  MouseMove($PosXInicial,$PosYInicial,1)
	  Sleep(300)

	  ClipPut("")

	  Send("!q")
	  Sleep(500)

	  MouseMove($PosXFinal,$PosYFinal,1)
	  Sleep(300)

	  Send("!q")
	  Sleep(2000)

	  Local $sCaptura = ClipGet()

	  Send("{ESC}")

	  if StringLen($sCaptura) > 0 Then
		 if $iDebug = 1 Then
			MsgBox(0,"Debug Valida",$sCaptura)
		 EndIf
		 ProcessClose("Capture2Text.exe")
		 Return True
	  EndIf

	  Sleep($TempoEsperaTentativa)
   Next

   ProcessClose("Capture2Text.exe")
   Return False
EndFunc

func _BuscarTextoOcr($PosXInicial,$PosYInicial,$PosXFinal,$PosYFinal,$iDebug = 0)

   Run($CaminhoOCR & "Capture2Text.exe")

   Sleep(100)

   For $i = 1 To 10

	  Send("{ESC}")
	  ClipPut("")
	  MouseMove($PosXInicial,$PosYInicial,1)
	  Sleep(300)

	  Send("!q")
	  Sleep(500)

	  MouseMove($PosXFinal,$PosYFinal,1)
	  Sleep(300)

	  Send("!q")
	  Sleep(2000)

	  MsgBox(0,"", "Apertou")

	  Global $sCaptura = ClipGet()

	  Send("{ESC}")

	  if StringLen($sCaptura) > 0 Then
		 if $iDebug = 1 Then
			MsgBox(0,"Debug Valida",$sCaptura)
		 EndIf

		 ProcessClose("Capture2Text.exe")
		 Return $sCaptura
	  EndIf

   Next

   ProcessClose("Capture2Text.exe")
   Return Null

EndFunc

Func _ValidaConta($PosXInicial,$PosYInicial,$PosXFinal,$PosYFinal,$Conta1,$NumTentativas,$QtdNumIguais)

   Run($CaminhoOCR & "Capture2Text.exe")

   Sleep(2000)

   For $i = 1 to $NumTentativas

	  Send("{ESC}")

	  MouseMove($PosXInicial,$PosYInicial,1)
	  Sleep(100)

	  ClipPut("")

	  Send("!q")
	  Sleep(100)

	  MouseMove($PosXFinal,$PosYFinal,1)
	  Sleep(100)

	  Send("!q")
	  Sleep(1000)

	  Global $sCaptura = ClipGet()

	  Send("{ESC}")

     Local $Conta2 = $sCaptura
     Local $igual = 0
    $Conta2 = StringReplace($Conta2," ","")

	  For $x = 1 To StringLen($Conta1)

		If StringMid($Conta1, $x, 1) =  StringMid($Conta2, $x, 1) Then

			   $igual = $igual + 1

		 EndIf

	  Next

	  If $igual >= $QtdNumIguais Then
			ExitLoop
	  EndIf

;~ 	  if $ExecMaximizar =  1 Then
;~ 		Send("!{SPACE}x")
;~ 	  EndIf
;~ 	  Sleep($TempoEsperaTentativa)
   Next

   If $igual >= $QtdNumIguais Then
	  Return True
   EndIf

   ProcessClose("Capture2Text.exe")
   Return False
EndFunc


Func _ValidaValor($PosXInicial,$PosYInicial,$PosXFinal,$PosYFinal,$NumTentativas)

   Run($CaminhoOCR & "Capture2Text.exe")

   Sleep(2000)

   For $i = 1 to $NumTentativas

	  Send("{ESC}")

	  MouseMove($PosXInicial,$PosYInicial,1)
	  Sleep(100)

	  ClipPut("")

	  Send("!q")
	  Sleep(100)

	  MouseMove($PosXFinal,$PosYFinal,1)
	  Sleep(100)

	  Send("!q")
	  Sleep(1000)

	  Global $sCaptura = ClipGet()

	  Send("{ESC}")

	  If $sCaptura <> "" Or $sCaptura <> null Then
		 ProcessClose("Capture2Text.exe")
		 Return $sCaptura
	  EndIf

   Next

   ProcessClose("Capture2Text.exe")
   Return ""
EndFunc

Func _ValidaConta2($PosXInicial,$PosYInicial,$PosXFinal,$PosYFinal)

   Run($CaminhoOCR & "Capture2Text.exe")

   Sleep(2000)

   For $i = 1 to $NumTentativas

	  Send("{ESC}")

	  MouseMove($PosXInicial,$PosYInicial,1)
	  Sleep(100)

	  ClipPut("")

	  Send("!q")
	  Sleep(100)

	  MouseMove($PosXFinal,$PosYFinal,1)
	  Sleep(100)

	  Send("!q")
	  Sleep(1000)

	  Global $sCaptura = ClipGet()

	  Send("{ESC}")

	  If $sCaptura <> "" Or $sCaptura <> null Then
		 ProcessClose("Capture2Text.exe")
		 Return $sCaptura
	  EndIf

   Next

   ProcessClose("Capture2Text.exe")
   Return ""
EndFunc

Func _Teste($PosXInicial,$PosYInicial,$PosXFinal,$PosYFinal,$TempoEsperaTentativa)

	   Run($CaminhoOCR & "Capture2Text.exe")

	  Sleep(2000)

	  Do

	  Send("{ESC}")

	  MouseMove($PosXInicial,$PosYInicial,1)
	  Sleep(100)

	  ClipPut("")

	  Send("!q")
	  Sleep(100)

	  MouseMove($PosXFinal,$PosYFinal,1)
	  Sleep(100)

	  Send("!q")
	  Sleep(1000)

	  Global $sCaptura = ClipGet()

	  Send("{ESC}")

	  If $sCaptura <> "" Or $sCaptura <> null Then
		 ProcessClose("Capture2Text.exe")
		 Return $sCaptura
	  EndIf

	  Until $sCaptura <> ""

EndFunc


Func _Teste2($PosXInicial,$PosYInicial,$PosXFinal,$PosYFinal,$TempoEsperaTentativa,$MesAnoInformado)

$PosYInicial = 340
$PosYFinal = 356
GLobal $PosYValInicio = 341
Global $PosYValFinal = 356
Local $ContaData
;Global Const $hexaBarraGrid  = "D6D6CE"

;Global $iColor  = PixelGetColor(1316,578)
;Global $HexaColor = Hex($iColor, 6)

;~ Local $PosXDataInicial = 63
;~ Local $PosYDataInicial = 339
;~ Local $PosXDataFinal = 104
;~ Local $PosYDataFinal = 354

	  Run($CaminhoOCR & "Capture2Text.exe")

	  Sleep(2000)

	  Do

		 Send("{ESC}")

		 MouseMove($PosXInicial,$PosYInicial,1)
		 Sleep(100)

		 ClipPut("")

		 Send("!q")
		 Sleep(100)

		 MouseMove($PosXFinal,$PosYFinal,1)
		 Sleep(100)

		 Send("!q")
		 Sleep(1000)

		 Global $sCaptura = ClipGet()

;~ 		 MouseMove($PosXDataInicial,$PosYDataInicial,1)
;~ 		 Sleep(100)

;~ 		 ClipPut("")

;~ 		 Send("!q")
;~ 		 Sleep(100)

;~ 		 MouseMove($PosXDataFinal,$PosYDataFinal,1)
;~ 		 Sleep(100)

;~ 		 Send("!q")
;~ 		 Sleep(1000)

;~ 		 Global $Data = ClipGet()

;~ 		 MsgBox(0,"",$Data)
;~ 		 Exit

		 Send("{ESC}")

		 If $sCaptura <> "" Then
				_MARCACAO_CORINGA(1314,490,$sApp[$eSleepTime])
				_MARCACAO_CORINGA($PosXInicial,$PosYInicial,$sApp[$eSleepTime])
			   Return $sCaptura
			   ProcessClose("Capture2Text.exe")
			Else
			   Return $sCaptura
			   ProcessClose("Capture2Text.exe")
		 EndIf

	  Until $sCaptura = ""

EndFunc

Func _Teste3($PosXInicial,$PosYInicial,$PosXFinal,$PosYFinal,$TempoEsperaTentativa)

$PosYInicial = 340
$PosYFinal = 356

	   Run($CaminhoOCR & "Capture2Text.exe")

	  Sleep(2000)

	  Do

	  Send("{ESC}")

	  MouseMove($PosXInicial,$PosYInicial,1)
	  Sleep(100)

	  ClipPut("")

	  Send("!q")
	  Sleep(100)

	  MouseMove($PosXFinal,$PosYFinal,1)
	  Sleep(100)

	  Send("!q")
	  Sleep(1000)

	  Global $sCaptura = ClipGet()

	  Send("{ESC}")

	  If $sCaptura <> "" Then
			_MARCACAO_CORINGA(1314,490,$sApp[$eSleepTime])
			_MARCACAO_CORINGA($PosXInicial,$PosYInicial,$sApp[$eSleepTime])
			Return $sCaptura
			ProcessClose("Capture2Text.exe")
		 Else
			$PosYInicial = $PosYInicial + 20
			$PosYFinal = $PosYFinal + 20
	  EndIf

	  Until $sCaptura = ""

EndFunc

Func _ValidaData($PosXInicial,$PosYInicial,$PosXFinal,$PosYFinal,$TempoEsperaTentativa)

Local $PosXDataInicial = 63
Local $PosYDataInicial = 339
Local $PosXDataFinal = 104
Local $PosYDataFinal = 354

 Run($CaminhoOCR & "Capture2Text.exe")

 Sleep(2000)

 Do

		 MouseMove($PosXDataInicial,$PosYDataInicial,1)
		 Sleep(100)

		 ClipPut("")

		 Send("!q")
		 Sleep(100)

		 MouseMove($PosXDataFinal,$PosYDataFinal,1)
		 Sleep(100)

		 Send("!q")
		 Sleep(1000)

		 Global $Data = ClipGet()

		 MsgBox(0,"",$Data)
		 Exit

		 Send("{ESC}")

	  If $Data <> "" Then
			_MARCACAO_CORINGA(1314,490,$sApp[$eSleepTime])
			_MARCACAO_CORINGA($PosXInicial,$PosYInicial,$sApp[$eSleepTime])
			Return $sCaptura
			ProcessClose("Capture2Text.exe")
		 Else
			Return $sCaptura
			ProcessClose("Capture2Text.exe")
	  EndIf

 Until  $Data = $MesInformado



EndFunc
