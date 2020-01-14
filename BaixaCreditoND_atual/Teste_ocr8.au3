#include <File.au3>
;~ #include "bin\funcoes_UTIL_160830-1555.au3"
#include "bin\Capture2Text.au3"
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

;~ Local $teste = "            7797"

;~ Local $teste2 = StringReplace(StringReplace(StringReplace($teste,".",""),",",""),"	","")
Global Enum $eContaRetorno, $eDtPgtoRetorno, $eVlrRetorno
;~ ConsoleWrite($teste2)
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

Local $fileImgCheck = @ScriptDir & "\prints\PrintAtlysConverted_Teste.jpg"
Local $fileImgCheck2 = @ScriptDir & "\prints\PrintAtlys_Teste.jpg"

$iDays = _DateDiff("D", '2018-04-06', '2018-04-05')
MsgBox($MB_SYSTEMMODAL, "", $iDays)

;~ Local $sNewDate = _DateAdd('d', 5, '2018-04-4')
;~ MsgBox($MB_SYSTEMMODAL, "", "Today + 5 days: " & $sNewDate)

sleep(10000000000)


Local $listaImagem

	_GDIPlus_Startup()
		SaveBmp2JPG($fileImgCheck,$fileImgCheck2)
	_GDIPlus_Shutdown()

$listaImagem = _ChamaOcr(1,$fileImgCheck2) ;retorna uma lista com as contas que forem mostradas na tabela de pesquisa de contas.




_ArrayDisplay($listaImagem)

Return

 local $datastring = "20/03/2013 R$1.311,35  DINHEIRO  PAGAMENTO COM ERRO 21/03/2013" ;= sucesso
;~ local $datastring = "20/03/2013 R$ 2.302,30 DINHEIRO  PAGAMENTO COM ERRO 21/03/2013" ;= sucesso
;~ local $datastring = "20/03/2013  R$231,47   DINHEIRO  PAGAMENTO COM ERRO 21/03/2013" ;= sucesso
;~ local $datastring = "20/03/2013 R$2.197,34  DINHEIRO  PAGAMENTO COM ERRO 21/03/2013" ;= sucesso
;~ local $datastring = "20/03/2013  R$ 906,52  DINHEIRO  PAGAMENTO COM ERRO 21/03/2013" ;= sucesso
;~ local $datastring = "20/03/2013 R$ 3.373,30 DINHEIRO  PAGAMENTO COM ERRO 21/03/2013" ;= sucesso
;~ local $datastring = "19/03/2013  R$ 447,09  DINHEIRO  PAGAMENTO COM ERRO 21/03/2013" ;= sucesso

local $dd = StringLeft($datastring,2)
local $mm = StringRight(StringLeft($datastring,5),2)
local $yyyy = StringRight(StringLeft($datastring,10),4)
local $yyyymmdd =  $yyyy & $mm & $dd


local $procuras = StringInStr($datastring,"$")
local $stringlen = StringLen($datastring)
local $tratado = StringRight($datastring,$stringlen-$procuras)
local $procuradinheiro = StringInStr($tratado,"DINHEIRO")
local $valor = StringLeft($tratado,$procuradinheiro - 1)

Local $valor = StringLeft(StringRight($datastring,$stringlen-$procuras),$procuradinheiro)

;~ MsgBox(0,"",$procuras)
;~ MsgBox(0,"",$stringlen)
;~ MsgBox(0,"",$yyyymmdd)


;~ local $acertou = _validaPercentualAcerto("Teste", "testesas")

;~ MsgBox(0,"",$acertou)

sleep(10000000)


Local $xtlinha = StringInStr($teste,"/")
Local $xtlinha2 = StringInStr($teste2,"/")

MsgBox(0,"",$xtlinha)
MsgBox(0,"",$xtlinha2)

sleep(10000000)

; teste de chamada
local $ApplicationId = "TechForOps_TLF_BPO"
local $Password = "HyxdUzcN4oMVZSOsdCEcQKlK";
local $FilePath, $postData
local $toEncode, $token
local $txtRetorno, $erro

Global $aRetornoTela1[3]
Global Enum $eContaRetorno, $eDtPgtoRetorno, $eVlrRetorno


Local $buscarTela = _identificarTela($arrTela[$eAtlys],1000,5)
	$buscarTela = $arrTela[$eAtlys]
		Local $hWnd = WinWait($arrTela[$eAtlys], "", 10)
		Local $aClientSize = WinGetClientSize($hWnd)

		WinActivate($buscarTela)


Global Const $fileImg = @ScriptDir & "\prints\PrintAtlys.tif"
Global Const $fileImg2 = @ScriptDir & "\prints\PrintAtlysConverted.tif"

		 Local $iColorPx  = PixelGetColor(1263, 452)
		 if $iColorPx = 0x000000 Then
			 MouseClick("left",1263, 452)
				send("{Space}")
		 EndIf

		 sleep(1000)

 _ScreenCapture_Capture($fileImg, _
											 $sPesquisaPagamento[$eMapaPrintTLoteInicioX], _
											 $sPesquisaPagamento[$eMapaPrintTLoteInicioY], _
											 $sPesquisaPagamento[$eMapaPrintTLoteFimX]	 , _
											 $sPesquisaPagamento[$eMapaPrintTLoteFimY],False)

		 _GDIPlus_Startup()
		 SaveBmp2JPG($fileImg,$fileImg2)
		 _GDIPlus_Shutdown()


; converse appid e senha no Token
$toEncode = $ApplicationId & ":" & $Password
$token = _Base64Encode($toEncode)

; teste 1
$FilePath = @ScriptDir & "\prints\PrintAtlys.tif"
$postData = FileRead($FilePath)
MsgBox(0,"",$postData)
;~ ConsoleWrite($postData)

;~ Exit

ConsoleWrite("StringLen($txtRetorno)=" & StringLen($txtRetorno))

;$txtRetorno = StringRight($txtRetorno,StringLen($txtRetorno)-3)

if  _GetResultOcr($token, $postData, "PortugueseBrazilian", "txt", $txtRetorno, $erro) = false then
	ConsoleWrite("$erro=" & $erro & @CRLF)
Else
	ConsoleWrite("$txtRetorno=" & @CRLF & $txtRetorno & @CRLF)
;~ 	Local $procura = StringSplit(StringRegExp($txtRetorno,'\d+', 1),@CRLF)
;~     Local $procura = StringSplit(_ALLTRIM($txtRetorno),@CRLF)

;~    For $i = 1 to $procura[0]
;~ 	  if StringLen($procura[$i]) > 0 Then

;~ 		 Local $registro = _filtroCaracteres($procura[$i])
;~ 		 ConsoleWrite($registro & @CRLF)

;~ 		 $aRetornoTela1[$eContaRetorno] 	= StringLeft($registro,10)
;~ 		 $aRetornoTela1[$eDtPgtoRetorno] 	= StringMid($registro,11,10)
;~ 		 $aRetornoTela1[$eVlrRetorno] 		= StringMid($registro,21,StringLen($registro)-20)

;~ 	  EndIf
;~    Next



;~ _ArrayDisplay($procura)
endif

; teste 2
;~ $FilePath = "C:/Users/marcelo.zampereti/Downloads/Picture_010.tif";
;~ $postData = FileRead($FilePath)
;~ if  _GetResultOcr($token, $postData, "PortugueseBrazilian", "txt", $txtRetorno, $erro) = false then
;~ 	ConsoleWrite("$erro=" & $erro & @CRLF)
;~ Else
;~ 	ConsoleWrite("$txtRetorno=" & @CRLF & $txtRetorno & @CRLF)
;~ endif
; teste de chamada



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
	local $filtro = "1234567890/"
;~ 	local $codiB = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN"
	local $temp = $caract
	local $result = ""

		For $i = 1 To StringLen($caract)
			$p = StringInStr($filtro, StringMid($caract, $i, 1))
			If $p > 0 Then
				$result = $result & StringMid($caract, $i, 1)
			EndIf
		Next

	Return $result
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

	Local $hBitmap_Scaled = _GDIPlus_ImageResize($Bitmap, $iW + 50 , $iH + 50)


    DllStructSetData($tData, "Quality", $iQuality)
    _GDIPlus_ParamAdd($tParams, $GDIP_EPGQUALITY, 1, $GDIP_EPTLONG, $pData)
    ;If Not _GDIPlus_ImageSaveToFileEx($Bitmap, $sSave, $sCLSID, $pParams) Then Return SetError(2, 0, 0)
	If Not _GDIPlus_ImageSaveToFileEx($hBitmap_Scaled, $sSave, $sCLSID, $pParams) Then Return SetError(2, 0, 0)
    Return True
 EndFunc

Func _ChamaOcr2($tipoDado,$fileImage)

   Local $txtRetorno
   Local $erro
   Local $aRetornoTela1[7][6]


   ;Tipo Dado 1 = Grid Busca de Pagamento
   $postData = FileRead($fileImage)
;~    MsgBox(0,"",$postData)

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
			For $i = 0 to UBound($procura) - 1
				ConsoleWrite("Valor $procura[" & $i & "] = " & $procura[$i] & @CRLF )
				$ilinha = StringInStr($procura[$i],"PAGAMENTO COM ERRO")

				if $ilinha = 0 Then
;~ 					$validtens[$j] = $procura[$i]
;~ 					$j=$j+1

				Else

					local $ibarra = StringInStr($procura[$i],"/")
					;ConsoleWrite("Valor $ibarra = " & $ibarra & @CRLF )

					If $ibarra > 5 Then ;Limpar caracteres estranhos
						$procura[$i] = StringRight($procura[$i],StringLen($procura[$i])-3)
					EndIf

					ConsoleWrite("Valor $j = " & $j & @CRLF )
					$validtens[$j] = $procura[$i]



					ReDim $validtens[UBound($validtens) + 1]
					$j=$j+1

				EndIf

			Next

;~ 				Next

				ConsoleWrite("Fim OCR" & @CRLF )
				Return $validtens

		 EndIf
		 Return False
	  EndIf
   Endif

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