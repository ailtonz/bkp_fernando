#include "WinHttp.au3" ; Necessário para funcionamento da GetResult

Func _GetResultOcr($token, $postData, $language, $exportFormat, byref $txtRetorno, byref $erro)
; #FUNCTION# ;===============================================================================
; Name...........: _GetResultOcr
; Description ...: Processa reconhecimento de imagem na Nuvem, ABBYY (https://ocrsdk.com/documentation/apireference/processImage/).
; Syntax.........: _GetResultOcr($token, $postData, $language, $exportFormat, byref $txtRetorno, byref $erro)
; Parameters ....: $token - Token composto pelo comando "$token=_Base64Encode($ApplicationId & ":" & $Password)" , cadastrado no ABBYY.
;                  $postData - Conteúdo do arquivo de imagem, usar comando "$postData = FileRead($FilePath)".
;                  $language - linguagem de reconhecimento usada pelo ABBYY.
;                  $exportFormat - tipo de retorno usada pelo ABBYY.
;                  $txtRetorno - variável que deverá retornar o conteúdo reconhecido pelo ABBYY.
;                  $erro - em caso de erro retronará o status code e comentário de erro.
; Return values .: Success - true
;                  |$txtRetorno : preenchida com reconhecimento do OCR ABBYY
;                  Failure - false
;                  |$erro : preenchida com comentário de erro
; Author ........: Marcelo Zampereti
; Related .......: #include "WinHttp.au3"
; Link ..........: https://ocrsdk.com/documentation/apireference/processImage/
; #include "WinHttp.au3" ; Necessário para funcionamento da GetResult
;============================================================================================

	local $apiAddess
	local $oHTTP, $oReceived, $oStatusCode, $id, $oHead
	local $status, $result, $resultUrl, $oXML
	local $cc = 0
	$txtRetorno=""
	$erro=""

	ConsoleWrite(@CRLF & "POST"& @CRLF)
	$oXML=ObjCreate("Microsoft.XMLDOM")

	;local $url = "http://cloud.ocrsdk.com/processImage?language=" & $language & " exportFormat=" & $exportFormat
	$apiAddess = "http://cloud.ocrsdk.com/processImage?language=" & $language & "&exportFormat=" & $exportFormat
	ConsoleWrite("$apiAddess="  & $apiAddess & @CRLF)

	; Envia Solicitação de OCR  "POST"
	$oHTTP = ObjCreate("winhttp.winhttprequest.5.1")
	$oHTTP.Open("POST", $apiAddess, False)
	$oHTTP.SetRequestHeader("Authorization", "Basic " & $token )
	$oHTTP.SetRequestHeader("Content-Disposition", "form-data")
	$oHTTP.SetRequestHeader("Content-Type", "image/jpeg")
	$oHTTP.Send($postData)
	ConsoleWrite("$oHTTP.ResponseText=" & $oHTTP.ResponseText & @CRLF)

	if $oHTTP.Status <> 200 Then
		$erro = "Erro #01, POST , $oHTTP.Status=" & $oHTTP.Status
		return false
	Else
		;Processou com sucesso o POST
		; Recupera ID do reponse
		$oXML.loadxml($oHTTP.ResponseText) ; load string
		$result = $oXML.selectSingleNode( '//response/task' )
		$id = $result.getAttributeNode( "id" )
		$status = $result.getAttributeNode( "status" )

		if $status.text = "Queued" and $status.text = "Completed"  then
			$erro = "Erro #02, POST, $status.text=" & $status.text
			$erro = $oHTTP.Status
			return false
		else
			$cc = 0
			DO
				ConsoleWrite(@CRLF & "GET"& @CRLF)

				; Consulta status do processamento "GET"
				local $urlGet = "http://cloud.ocrsdk.com/getTaskStatus?taskId=" & $id.text
				ConsoleWrite("$urlGet="  & $urlGet & @CRLF)

				$oHTTP.Open("GET", $urlGet, False)
				$oHTTP.SetRequestHeader("Authorization", "Basic " & $token )
				$oHTTP.Send()

				if $oHTTP.Status <> 200 Then
					$erro = "Erro #03, GET, $oHTTP.Status=" & $oHTTP.Status
					return false
				Else
                    ;Processou com sucesso o GET
					ConsoleWrite("$oHTTP.ResponseText=" & $oHTTP.ResponseText & @CRLF)

					; Recupera URL do reponse
					$oXML.loadxml($oHTTP.ResponseText) ; load string
					$result = $oXML.selectSingleNode( '//response/task' )
					$id = $result.getAttributeNode( "id" )
					$status = $result.getAttributeNode( "status" )
					ConsoleWrite(@CRLF)

					if $status.text="Completed" Then

						$resultUrl = $result.getAttributeNode( "resultUrl" )
						ConsoleWrite(@CRLF & "GET ARQUIVO TXT"& @CRLF)
						;ConsoleWrite ( "Getting value from XML file." & @CRLF)
						ConsoleWrite ( $resultUrl.xml & @CRLF)
						;ConsoleWrite ( "End ofs XML file." & @CRLF)

						; Busca arquivo TXT "GET"
						$oHTTP.Open("get", $resultUrl.text, False)
						$oHTTP.Send()

						if $oHTTP.Status <> 200 Then
							$erro = "Erro #04, GET TXT, $oHTTP.Status=" & $oHTTP.Status
							return false
						Else
							;Processou com sucesso o GET

							; Recupera Arquivo TXT no response
							$txtRetorno = $oHTTP.ResponseText
							ConsoleWrite ("$txtRetorno = " & $txtRetorno)
							; Processado com sucesso
							return true
							;ExitLoop
						endif
					Else
						if $cc >= 10 Then
							$erro = "Erro #05, GET , $status.text=" & $status.text
							return false
						else
							$cc = $cc + 1
							sleep(1000)
						endif
					EndIf

				endif

			UNTIL $status="Completed"

		endif

	endif

EndFunc

