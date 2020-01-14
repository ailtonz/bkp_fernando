#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <GUIConstantsEx.au3> ; constante para GUI eventos
#include <AutoItConstants.au3>
#include <EditConstants.au3>
#include <File.au3>
#include <IE.au3>
#include <GuiTreeView.au3>
#include <ScreenCapture.au3>
#include <SQL_Connect.au3>
#include <StringConstants.au3>
#include "PosVariaveis.au3"
#include "ApiWebService.au3"

;Quando pressionar a tecla 'Esc' sai da rotina
HotKeySet("{End}", "CaptureEsc")


Global $jsonToSendSTR = '{"FileName": "Nome_Arquivo","FileBase64": "YXJxdWl2byBkZSB0ZXN0ZQ=="}'

Global Enum $eLocal,$eDev,$eQA,$ePrd
Global Enum $eCnnLocal,$eCnnAzure
Global Enum $eUserLocal,$eUserDev,$eUserQA,$eUserPrd

Global $aHttpServ[4]
	   $aHttpServ[$eLocal]  = "http://localhost:59911/"
	   $aHttpServ[$eDev]	= "http://telefonicabpo-juridico-development.azurewebsites.net/"
	   $aHttpServ[$eQA]		= "http://telefonicabpo-juridico-quality.azurewebsites.net/"
	   $aHttpServ[$ePrd]	= "http://telefonicabpo-juridico.azurewebsites.net/"

Global $aStringConn[2]
	   $aStringConn[$eCnnLocal]  = "localhost"
	   $aStringConn[$eCnnAzure]	= "juridicoserver.database.windows.net"

Global $aUserConn[4][4]

	   ;Conexao Local Sem Credencias de Acesso
	   $aUserConn[$eUserLocal][0] = $aStringConn[$eCnnLocal]
	   $aUserConn[$eUserLocal][1] = "DB_ROBO_JURIDICO"
	   $aUserConn[$eUserLocal][2] = Null
	   $aUserConn[$eUserLocal][3] = Null

	   ;Conexao Banco Dev
	   $aUserConn[$eUserDev][0] = $aStringConn[$eCnnAzure]
	   $aUserConn[$eUserDev][1] = "JuridicoDEV"
	   $aUserConn[$eUserDev][2] = "app_Juridico_DEV"
	   $aUserConn[$eUserDev][3]  = "_eb!KACas2eg5ve6"

	   ;Conexao Banco Hml
	   $aUserConn[$eUserQA][0] = $aStringConn[$eCnnAzure]
	   $aUserConn[$eUserQA][1] = "JuridicoHML"
	   $aUserConn[$eUserQA][2] = "app_Juridico_HML"
	   $aUserConn[$eUserQA][3]  = "prepu4pus5_#aWe@"

	   ;Conexao Banco Prd
	   $aUserConn[$eUserPrd][0] = $aStringConn[$eCnnAzure]
	   $aUserConn[$eUserPrd][1] = "JuridicoPRD"
	   $aUserConn[$eUserPrd][2] = "app_Juridico_PRD"
	   $aUserConn[$eUserPrd][3]  = "baDr=sUDRE45eka!"

;Alterar a Conexao Quando Mudar o processo de ambiente;
Global $sHttpServ = $aHttpServ[$eQA]

;Alterar a Conexao Quando Mudar o processo de ambiente;
Global $sConnRobo 	  = $aUserConn[$eUserQA][0]
Global $sDbNameRobo	  = $aUserConn[$eUserQA][1]
Global $sUserRobo 	  = $aUserConn[$eUserQA][2]
Global $sPwdRobo  	  = $aUserConn[$eUserQA][3]

Global $token = "Basic bC5hLmp1bmlvcjpMakAyODE1MDI="

Global Enum $eWinLogon,$eAtlys,$eNgin,$eVivoNet,$eAtlysSolution,$eBloqueioAtlys,$eCSO,$eAtlysErro
Global $sDiretorio  = @ScriptDir & "\Includes\Config\"
Global $sConnectionCenter = "C:\Program Files (x86)\Citrix\ICA Client\concentr.exe"
Global $arrTela[8]
	   $arrTela[$eWinLogon] 		= "Windows Logon - \\Remote"
	   $arrTela[$eAtlys] 			= "Atlys Global Solution - \\Remote"
	   $arrTela[$eNgin] 			= "inVIVO"
	   $arrTela[$eVivoNet] 			= "Vivo Net - Windows Internet Explorer - \\Remote"
	   $arrTela[$eAtlysSolution] 	= "Logon do Atlys Global Solution - \\Remote"
	   $arrTela[$eBloqueioAtlys] 	= "Bloqueio da Solução Global Atlys - \\Remote"
	   $arrTela[$eCSO]				= "pw3270:"
	   $arrTela[$eAtlysErro]		= "Atlys ERROR"

Global $arrNomesistema[7]
	   $arrNomesistema[$eWinLogon]  	= "Logon"
	   $arrNomesistema[$eAtlys] 		= "Atlys"
	   $arrNomesistema[$eNgin] 			= "Ngin"
	   $arrNomesistema[$eVivoNet] 		= "VivoNet"
	   $arrNomesistema[$eAtlysSolution]	= "Atlys"
	   $arrNomesistema[$eCSO] 			= "CSO"

Global Enum $eSistemaAtlys = 1, $eSistemaNgin = 2, $eSistemaVivoNet = 3, $eSistemaCitrix =4,$eSistemaCSO = 5
Global Enum $eProcessoBuscaTerminais 	= 1, _
			$eProcessoAtlysContas 		= 2, _
			$eProcessoNgin		 		= 3, _
			$eProcessoVivoNet	 		= 4, _
			$eProcessoEnvioLinte 		= 5, _
			$eProcessoAtlysTerminais	= 6

Global Enum $eRotaSistemaUsuario,$eRotaInativarUsuario,$eRotaFileBytes,$eRotaStatusProcesso,$eRotaEnvioEmailFtpSucesso
Global $arrApiRota[5]

	  $arrApiRota[$eRotaSistemaUsuario]  		= "api/sistemaUsuario/listar/"

	  ;Chamada da Rota Necessita de Paramentro do Id do Usuario que sera inativado
	  $arrApiRota[$eRotaInativarUsuario] 		= "api/sistemaUsuario/inativar/"

	  ;Envio Azure
	  $arrApiRota[$eRotaFileBytes]				= "api/file/bytes/"

	  $arrApiRota[$eRotaStatusProcesso] 		= "api/solicitacao/processoStatus"

	  ;Chamada da Api necessita de parametro do Id da solicitacao
	  $arrApiRota[$eRotaEnvioEmailFtpSucesso] 	= "api/ftpSucesso/enviarEmail/"


Global $CaminhoOCR  = @ScriptDir & "\Includes\Config\Capture2Text_v4.4.0_64bit\"
;~ Global $CaminhoOCR2 = @ScriptDir & "\Config\Capture2Text_v4.4.0_64bit\"

;Enum para o Status do Processamento
Global Enum $eStProcAguardando  =1, _
			$eStProcProcessando =2, _
			$eStProcFinalizado  =3, _
			$eStProcErro 		=4


Global $sDiretorioImagensProcesso = "\\Tlf-prd-rpa001\e\"

Global $ipAdress = ""
Global Const $pwsVpn = "vpnaccenture"

Global enum $eFtpServer,$eFtpUser,$eFtpPsw,$eFtpDir
Global $arrayCredFtp[4]
	   $arrayCredFtp[$eFtpServer] 	= "52.44.30.201:22"
	   $arrayCredFtp[$eFtpUser] 	= "telefonica"
	   $arrayCredFtp[$eFtpPsw] 		= "Au6qEdV7NRN6ZFpe2kKZYM92uPv7tZ"
	   $arrayCredFtp[$eFtpDir] 		= "repository/files"

Global $idVmProcessamento

Func _BuscarCredencialCitrix()

   Global $CredCitrix = _RetornaUsuario($eSistemaCitrix,$idVmProcessamento)
   if $CredCitrix = -1 Then
	  ;Criar Log De Execucao Logs
	  Return False
   EndIf

   Global $UserCitrix 		= $CredCitrix[0][0]
   Global $PassCitrix 		= $CredCitrix[1][0]
   Global $IdUsuarioCitrix 	= $CredCitrix[2][0]

   If StringLen($UserCitrix) = 0 Then
	  ;Criar Log De Execucao Logs
	  Return False
   EndIf

   Return True
EndFunc

Func _VerificaCitrixLogado()
   ;=======================================================
   ;Verifica se o citrix possui estancias do sistema aberto
   ;=======================================================
   ShellExecuteWait($sConnectionCenter)
   Sleep(2000)

   Local $hWnd = ControlGetHandle("[CLASS:#32770;TITLE:Citrix Connection Center]","","[CLASS:SysTreeView32; INSTANCE:1]")
   Sleep(2000)

   Local $iCount = _GUICtrlTreeView_GetCount ($hWnd)

   if $iCount > 0 Then
	  Sleep(1000)
	  Send("!C")
	  Sleep(1000)
	  Return True
   Else
	  Send("!C")
	  Sleep(1000)
	  Return False
   EndIf
EndFunc

Func _iniciarSistema($idSistema,$citrixUser,$citrixPass,$idUser)

   ;Informacao de Mensagem de Erro senha Citrix
   ;===========================================
   Local $arrayMsgSenha[2]
		 $arrayMsgSenha[0] = "incorrect"
		 $arrayMsgSenha[1] = "locked"
   ;===========================================

   ;Retorno 1  - Sistema Logado
   ;Retorno 2  - Tela Login
   ;Retorno -1 - Erro


   ;Funcao tenta realizar o Login do Citrix a partir de um arquivo de configuracao ica.
   ;Esse arquivo deve estar na pasta
   For $i = 1 TO 10

	  Local $nomeArquivoICA  = "iniciar" & $arrNomesistema[$idSistema] & ".ica"
	  Local $enderecoArquivo = $sDiretorio & $nomeArquivoICA
	  ShellExecuteWait($enderecoArquivo)

	  Local $buscarTela = _identificarTela($arrTela[$eWinLogon],4000)
	  if $buscarTela = $arrTela[$eWinLogon] Then

		 WinActivate($buscarTela)
		 Send("{ENTER}")

		 Sleep(2000)

		 WinActivate($buscarTela)
		 Send($citrixUser)
		 Sleep(100)

		 WinActivate($buscarTela)
		 Send("{TAB}")
		 Sleep(100)

		 WinActivate($buscarTela)
		 Send($citrixPass)
		 Sleep(100)

		 WinActivate($buscarTela)
		 Send("{TAB}")
		 Sleep(100)

		 WinActivate($buscarTela)
		 Send("{ENTER}")
		 Sleep(7500)

;~ 		 Local $fileCitrix = @ScriptDir & "\Includes\tmp\screenCitrix.jpg"
;~ 		 Local $fileTitle  = "screenCitrix.jpg - Visualizador de Fotos do Windows"
;~ 		 MouseMove(0,0)
;~ 		 _ScreenCapture_Capture($fileCitrix)


		 $idBuscaTela = 1
		 Local $buscarTela = _identificarTela($arrTela[$idSistema],2000,10,$idBuscaTela)
		 if StringInStr($buscarTela,$arrTela[$idSistema]) > 0 Then
			WinActivate($buscarTela)
;~ 			FileDelete($fileCitrix)
			Return 2
		 EndIf

;~ 		 ShellExecute($fileCitrix)
		 WinActivate($buscarTela)
		 if _ValidaTextoTela($arrayS1ListaPosicao[$eS4T0MsgSenhaErradaInicioX]		 , _
									  $arrayS1ListaPosicao[$eS4T0MsgSenhaErradaInicioY], _
									  $arrayS1ListaPosicao[$eS4T0MsgSenhaErradaFimX]   , _
									  $arrayS1ListaPosicao[$eS4T0MsgSenhaErradaFimY]   , _
									  $arrayMsgSenha,3,1000) Then

			apiCall($sHttpServ,$arrApiRota[$eRotaInativarUsuario] & $idUser,"PUT","",$token,$eTech,0,"","")
			ConsoleWrite("---------Senha Citrix Incorreta " & @CRLF)
			WinClose($fileTitle)
			FileDelete($fileCitrix)
			Return -1
		 EndIf

		 Local $buscarTela = _identificarTela($arrTela[$idSistema],2000,10,$idBuscaTela)
		 if StringInStr($buscarTela,$arrTela[$idSistema]) > 0  Then
			WinActivate($buscarTela)
;~ 			FileDelete($fileCitrix)
			Return 2
		 EndIf
	  EndIf
   Next
   Return -1
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

Func _RetornaUsuario($idSistema,$idVmRequisicao)

   Local  $jsonToSendSTR = '{"idSistema":' & $idSistema & ',"idVmRequisicao":' & $idVmRequisicao & '}'

   Local $arrList = apiCall($sHttpServ,$arrApiRota[$eRotaSistemaUsuario],"GET",$jsonToSendSTR,$token,$eTech,1,"","")
   if $arrList = 0 Then
	  Return -1
   EndIf

   Local $pass	  = $arrList.Data
   Local $arrRet[3][1]

   $arrRet[0][0] = $pass.item(0).Usuario
   $arrRet[1][0] = $pass.item(0).Senha
   $arrRet[2][0] = $pass.item(0).Id

   Return $arrRet
EndFunc

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

	  Local $sCaptura = ClipGet()

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

Func CaptureEsc()

   Switch @HotKeyPressed
	  Case "{END}"
		 Local $iRetorno = MsgBox(4,"Pause","Deseja parar o script?")
		 If $iRetorno = 6 Then
			Exit
		 EndIf
   EndSwitch

EndFunc

Func _Connect()
;~    $__g_oTemplateCOMErrorHandler = ObjEvent("AutoIt.Error", "_erroLogConn")

   Local $Server   = $sConnRobo
   Local $DataBase = $sDbNameRobo
   Local $UserCnn  = $sUserRobo
   Local $PassCnn  = $sPwdRobo

   if $Server = "localhost" Then
	  Local $Conn	   = _SQLConnect($Server,$DataBase)
   Else
	  Local $Conn	   = _SQLConnect($Server,$DataBase,1,$UserCnn,$PassCnn,"{SQL Server}")
   EndIf

   if @error Then
	  Return 1
   EndIf

   Return $Conn
EndFunc

Func _AtualizarLogProcessamento($idLogProcessamento,$qtdTentativas,$IdStatusProcessamento, $obs = "",$idSolicitacao = 0)

   Local $sSql = "UPDATE [dbo].[TB_LOG_PROCESSAMENTO] SET "
   Local $cnn = _Connect()
   If $cnn.State = 1 Then

	  Switch $IdStatusProcessamento
		 Case 2
			if $qtdTentativas = 1 Then
			   $sSql =  $sSql & "[DT_INICIO] = GETDATE(),"

			   Local $strEnvio = '{"id" : ' & $idSolicitacao & ',"idProcessoStatus" :4}'
			   apiCall($sHttpServ,$arrApiRota[$eRotaStatusProcesso],"PUT",$strEnvio,$token,$eTech,1,"","")
			EndIf
		 Case 3
			$sSql =  $sSql & "[DT_FIM] = GETDATE(),"
			if StringLen($obs) > 0 Then
			   $sSql =  $sSql & "[OBS] = '" & $obs & "',"
			EndIf
	  EndSwitch

	  $sSql =  $sSql & "[QTD_TENTATIVAS] = " & $qtdTentativas & ","
	  $sSql =  $sSql & "[ID_STATUS_PROCESSAMENTO] = " & $IdStatusProcessamento
	  $sSql =  $sSql & " WHERE ID = " & $idLogProcessamento
	  $cnn.Execute($sSql)
   EndIf

EndFunc

Func _AtualizarLogRota($idRota,$QtdProcessos,$idSolicitacao,$sCiv)

   ;Baixar Rota
   Local $sSql 			= "UPDATE [dbo].[TB_ROTAS] SET [STATUS] = 1 , [DT_FIM] = GETDATE() WHERE ID = " & $idRota

   ;Identificar a quantidade processos finalizados de acordo com a rota
   Local $sSqlCount 	= "Select Count(*) FROM [dbo].[TB_LOG_PROCESSAMENTO] Where Id_Rota = " & $idRota & " AND [ID_STATUS_PROCESSAMENTO] = " & $eStProcFinalizado

   ;Marcar Inicio do Processamento
   Local $sSqlInicio 	= "UPDATE [dbo].[TB_ROTAS] SET [DT_INICIO] = GETDATE() WHERE ID = " & $idRota

   ;Nome da Procedure de Distribuicao
   Local $sProcFila		= "[dbo].[Distribuir_Fila]"

   ;Verificar se Processo Finalizado
   Local $sSqlContarRotaFinalizada = "Select Count(*) From [dbo].[TB_ROTAS] Where [ID_SOLICITACAO] = " & $idSolicitacao & " AND STATUS = 0"

   Local $cnn 	 = _Connect()

   Local $rs  	 = ObjCreate("ADODB.Recordset")
   If $cnn.State = 1 Then

	  $rs.open ($sSqlCount,$cnn)
	  Local $iQtd = $rs(0).Value
	  $rs.close

	  ;Se a quantidade de processos com status finalizado for igual ao total de processos mapeados por rota
	  if $iQtd = $QtdProcessos Then

		 ;Finalizar Log Rota
		 $cnn.Execute($sSql)

		 ;Redistribuir Fila Processamento
		 _ExecProcedure($sProcFila)

		 Local $cnn2 = _Connect()

		 If $cnn2.State = 1 Then

			Local $rs2  = ObjCreate("ADODB.Recordset")
			$rs2.open ($sSqlContarRotaFinalizada,$cnn2)
			Local $iQtd = $rs2(0).Value
			$rs2.close

			If $iQtd = 0 Then
				Local $sCaminhoProcesso = $sDiretorioImagensProcesso & $sCiv

				_compactar7z($sCaminhoProcesso,$sCaminhoProcesso & ".7z")

				If FileExists($sCaminhoProcesso & ".7z") Then
				   _EnviaSftp($sCiv & ".7z",$idSolicitacao)
				EndIf
			EndIf

		 EndIf

		 Return True
	  Elseif $iQtd = 0 then
		 $cnn.Execute($sSqlInicio)
	  EndIf
   EndIf
   Return False
EndFunc

Func _ExecProcedure($sProcName)

   Local $cnn =  _Connect()
   If $cnn.State = 1 Then
	  Local $objCmd = ObjCreate("ADODB.Command")
			$objCmd.ActiveConnection = $cnn
			$objCmd.CommandText = $sProcName
			$objCmd.CommandType = 4

			; request execution
			$objCmd.Execute

	  $cnn.Close
   EndIf
EndFunc

Func _GeraIdImagem($idLogProcessamento)

   Local $cnn = _Connect()
   Local $rs  	= ObjCreate("ADODB.Recordset")

   Local $sSql = "INSERT INTO [dbo].[TB_LOG_PRINTS] (ID_LOG_PROCESSAMENTO,DT_CRIACAO)"
		 $sSql = $sSql & " VALUES (" & $idLogProcessamento & ",GETDATE())" & @CRLF

   Local $sSqlLog = "SELECT ID AS LastID FROM [dbo].[TB_LOG_PRINTS] WHERE ID = @@Identity  AND [ID_LOG_PROCESSAMENTO] = " & $idLogProcessamento

   If $cnn.State = 1 Then
	  $cnn.Execute($sSql)
	  $rs.open ($sSqlLog,$cnn)
	  Return StringFormat("%011i",$rs(0).Value)
   EndIf

EndFunc

Func _DeletePrintsExistentes($sDiretorio,$sArquivo,$ext = ".jpg")

   Local $arrayFiles = _FileListToArray($sDiretorio,$sArquivo & "*" & $ext,1)

   If IsArray($arrayFiles) Then
	  For $i = 1 To $arrayFiles[0]
		 FileDelete($sDiretorio & "\" & $arrayFiles[$i])
	  Next
   EndIf
EndFunc

Func _compactar7z($File,$fileCompact)

   Local $cmd = $sDiretorio & "\7z.exe" & " a -mx=9 " &  $fileCompact & " " & $File
   RunWait (@ComSpec & " /c " & $cmd, @ScriptDir)
EndFunc

func _Convert($Caminho)

   Global $sFileName = _FileName($Caminho)
   Global $sFullFile = $Caminho
   Global $sFilePathjson = $sDiretorioImagensProcesso & $sFileName & ".json"


   Local $file = FileOpen($sFullFile, 16)
   Local $Text = FileRead($file)

   $sConverted = _Base64Encode($Text)

   Local $oJSON = _OO_JSON_Init()
		 $jsObj = $oJSON.parse($jsonToSendSTR) ; JSON encode key value
		 $jsObj.FileName 	= $sFileName
		 $jsObj.FileBase64 = $sConverted
		 $sPostData = $jsObj.stringify()

   if FileExists($sFilePathjson) Then
	  FileDelete($sFilePathjson)
   EndIf

   $hFileOpen = FileOpen($sFilePathjson, $FO_APPEND + $FO_BINARY + $FO_CREATEPATH)
   FileWrite($hFileOpen, $sPostData)

   Return $sPostData
EndFunc

Func _FileName($FullName)
   Local $fileSplit = StringSplit($FullName,"\")
   Return $fileSplit[$fileSplit[0]]
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


	  Local $sCaptura = ClipGet()

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

func _connectVpn()

	Local $sVpnAdress 		= "C:\Program Files (x86)\Cisco Systems\VPN Client\vpngui.exe"
	Local $sVpnConected 	= "status: Connected | VPN Client - Version 5.0.07.0290"
	Local $sVpnNotConected 	= "status: Disconnected | VPN Client - Version 5.0.07.0290"

    ShellExecute($sVpnAdress)
    Sleep(3000)

	if WinExists($sVpnConected) Then

;~ 		 WinSetState($sVpnConected,"",@SW_MAXIMIZE )
;~ 		 $ipAdress = _BuscarTextoOcr(957,709,1027,726,1)

		WinActivate($sVpnNotConected)
		Send("^q")

		Sleep(500)
		Return True

	Else

		WinActivate($sVpnNotConected)

		Send("^o")
		Sleep(1000)

		Send($pwsVpn)
		Sleep(1000)

		Send("{ENTER}")
		Sleep(6000)

		ShellExecute($sVpnAdress)
		Sleep(1000)

		if WinExists($sVpnConected) Then
			Return True
		Else
			Return False
		EndIf

   EndIf

;~    if Ping($ipAdress) = 1 Then
;~ 	  Return True
;~    Else
;~ 	  Return False
;~    EndIf

EndFunc

Func _EnviaSftp($fileName,$idSolicitacao)


Local $fileUfg 		= @ScriptDir & "\Includes\Config\winscp\winscpCfg_" & $fileName & ".cfg"
Local $fileWinScp 	= @ScriptDir & "\Includes\Config\winscp\winscp.com"
Local $fileLog 		= $sDiretorioImagensProcesso & "LogFtp\" & StringReplace($fileName,".7z",".log")

ConsoleWrite($fileUfg & @CRLF)
ConsoleWrite($fileWinScp & @CRLF)
ConsoleWrite($fileLog & @CRLF)

Local $cmd = $fileWinScp & " /script=" & $fileUfg & ">> " & $fileLog
ConsoleWrite($cmd)

Local $strEnvio = '{"id" : ' & $idSolicitacao & ',"idProcessoStatus" :5}'

if FileExists($fileUfg) Then
	FileDelete($fileUfg)
EndIf


FileWriteLine($fileUfg,"option batch abort" & @CRLF)
FileWriteLine($fileUfg,"option confirm off" & @CRLF)
FileWriteLine($fileUfg,"open sftp://" & $arrayCredFtp[$eFtpUser]   & ":" & _
										$arrayCredFtp[$eFtpPsw]    & "@" & _
										$arrayCredFtp[$eFtpServer] & "/" & @CRLF)
FileWriteLine($fileUfg,"cd /" & $arrayCredFtp[$eFtpDir] & @CRLF)
FileWriteLine($fileUfg,"option transfer binary" & @CRLF)
FileWriteLine($fileUfg,"put " & $sDiretorioImagensProcesso & $fileName & @CRLF)
FileWriteLine($fileUfg,"close" & @CRLF)
FileWriteLine($fileUfg,"exit" & @CRLF)

RunWait (@ComSpec & " /c " & $cmd, @ScriptDir)

apiCall($sHttpServ,$arrApiRota[$eRotaEnvioEmailFtpSucesso] & $idSolicitacao,"PUT",Default,$token,$eTech,Default,"","")
apiCall($sHttpServ,$arrApiRota[$eRotaStatusProcesso],"PUT",$strEnvio,$token,$eTech,1,"","")

FileDelete($sDiretorioImagensProcesso & $fileName)

EndFunc

Func _erroConexaoAtlys()
	If WinExists($arrTela[$eAtlysErro]) Then

		WinActivate($arrTela[$eAtlysErro])
		Sleep(1000)

		Send("{Enter}")
		Sleep(20000)

		_KillAtlys()
		_KillNgin()
		_KillVivoNet()

		Return True
	EndIf
EndFunc