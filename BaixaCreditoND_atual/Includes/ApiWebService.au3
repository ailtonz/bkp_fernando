

Global Enum $eRedecorp,$eTech

;~ Local $arrList = apiCall($aHttpServ[$eLocal],"api/sistemaUsuario/listar/4","GET","",$token,$eTech,0,"","")

Func APICall($apiAddess, $apiController = "", $actionType = "", $postData = "", $token = "",$Rede = $eRedecorp,$EnvioLista = 0,$userCitrix = "",$pwdCitrix = "")

   Local $jsObj
   Local $jsObjToken
   Local $jsObjList
   Local $oJSON = _OO_JSON_Init()

   If ($postData) Then
	  if not $EnvioLista Then
		 $jsObj = $oJSON.parse($postData) ; JSON encode key value
	  EndIf
   EndIf

   _WinHttpSetTimeouts(0,0,0,0)

   ;Verifica Qual Rede esta sendo usuada
   ;==================================================================================================

   if $Rede = $eRedecorp Then
	  Local $hOpen = _WinHttpOpen(Default, $WINHTTP_ACCESS_TYPE_NAMED_PROXY, "proxy.redecorp.br:8080")

   Else
	  Local $hOpen = _WinHttpOpen(Default)
   EndIf

   Local $hConnect = _WinHttpConnect($hOpen, $apiAddess, $INTERNET_DEFAULT_HTTP_PORT)

   ; Make a request
   Local $hRequest = _WinHttpOpenRequest( _
		 $hConnect, _
         $actionType, _
		 $apiController, _
		 Default, _
		 Default, _
		 "application/octet-stream") ;, _

	if $Rede = $eRedecorp Then
		_WinHttpSetCredentials($hRequest, $WINHTTP_AUTH_TARGET_PROXY, $WINHTTP_AUTH_SCHEME_BASIC, "redecorp\" & $userCitrix, $pwdCitrix)
	EndIf
;~    _WinHttpSetCredentials($hRequest, $WINHTTP_AUTH_TARGET_PROXY, $WINHTTP_AUTH_SCHEME_BASIC, "0785242", "186473")
   _WinHttpAddRequestHeaders($hRequest, "Content-Type: application/octet-stream")
   _WinHttpAddRequestHeaders($hRequest, "Ocp-Apim-Subscription-Key",$token )
;~    _WinHttpAddRequestHeaders($hRequest, "Connection: keep-alive")

;~    If ($token <> "" ) Then
;~ 	  _WinHttpAddRequestHeaders($hRequest, "Authorization: " & $token)
;~    EndIf

   ; Send it
   If ($postData) Then
	  if not $EnvioLista Then
		 ConsoleWrite("Lpp NOK")
		 _WinHttpSendRequest($hRequest, -1, $jsObj.stringify())
	  Else
		 ConsoleWrite("Lpp OK-----")
		 ConsoleWrite($postData)
		 _WinHttpSendRequest($hRequest, -1, $postData)
	  EndIf
   Else
	  ConsoleWrite("Lpp NNOK")
	  _WinHttpSendRequest($hRequest, -1)
   EndIf

   ; Wait for the response
   _WinHttpReceiveResponse($hRequest)

	; Check if there is a response
   Local $sHeader, $sReturned
   If _WinHttpQueryDataAvailable($hRequest) Then
	  $sHeader = StringSplit(_WinHttpQueryHeaders($hRequest), " ")

	  Do
		 $sReturned &= _WinHttpReadData($hRequest)
	  Until @error

	  ; Print returned
	  ConsoleWrite("Debug: " & $sReturned & @CRLF)

	  if($sReturned) then
		 $jsObjToken = $oJSON.parse($sReturned)
		 Return $jsObjToken
	  Else
		 return "Ok"
	  EndIf
   Else

	  ConsoleWriteError($jsObjToken & @CRLF)
   EndIf

   ; Close handles
   _WinHttpCloseHandle($hRequest)
   _WinHttpCloseHandle($hConnect)
   _WinHttpCloseHandle($hOpen)
EndFunc   ;==>APICall