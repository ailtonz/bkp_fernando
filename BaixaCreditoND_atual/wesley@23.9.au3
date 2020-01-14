#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <GUIConstantsEx.au3> ; constante para GUI eventos
#include <AutoItConstants.au3>
#include <EditConstants.au3>
#include <File.au3>
#include <IE.au3>
#include <GuiTreeView.au3>
#include <StringConstants.au3>
#include "Includes\ApiWebService.au3"
#include <FileConstants.au3>


Local $Print = "C:\Users\wesley.valadao\Downloads\BaixaCreditoND\CONTAS ATYLS.jpg"
;Global $jsonToSendSTR = '{"FileName": "Nome_Arquivo","FileBase64": "YXJxdWl2byBkZSB0ZXN0ZQ=="}'

Local $file = FileOpen($Print, 16)
Local $Text = FileRead($file)

;$sConverted = _URIDecode($Text)
Local $sConverted = _Base64Encode($Text)

ConsoleWrite($sConverted)


;Alterar a Conexao Quando Mudar o processo de ambiente;
Global $sHttpServ = "https://brazilsouth.api.cognitive.microsoft.com"
Local $Teste =  "/vision/v1.0"

Global $token = "Basic 9b59945e11fe421e87d9dcd83a8c4d5d"

apiCall($sHttpServ,,"PUT","",$token,$eTech,0,"","")

Local  $jsonToSendSTR = '{"idSistema":' & $idSistema & ',"idVmRequisicao":' & $idVmRequisicao & '}'

Local $arrList = apiCall($sHttpServ,$Teste,"POST",$jsonToSendSTR,$token,$eTech,1,"","")
if $arrList = 0 Then
   Return -1
EndIf

Local $pass	  = $arrList.Data
Local $arrRet[3][1]

$arrRet[0][0] = $pass.item(0).Usuario
$arrRet[1][0] = $pass.item(0).Senha
$arrRet[2][0] = $pass.item(0).Id



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
Func _Base64Decode($Data)
	Local $Opcode = "0xC81000005356578365F800E8500000003EFFFFFF3F3435363738393A3B3C3DFFFFFF00FFFFFF000102030405060708090A0B0C0D0E0F10111213141516171819FFFFFFFFFFFF1A1B1C1D1E1F202122232425262728292A2B2C2D2E2F303132338F45F08B7D0C8B5D0831D2E9910000008365FC00837DFC047D548A034384C0750383EA033C3D75094A803B3D75014AB00084C0751A837DFC047D0D8B75FCC64435F400FF45FCEBED6A018F45F8EB1F3C2B72193C7A77150FB6F083EE2B0375F08A068B75FC884435F4FF45FCEBA68D75F4668B06C0E002C0EC0408E08807668B4601C0E004C0EC0208E08847018A4602C0E00624C00A46038847028D7F038D5203837DF8000F8465FFFFFF89D05F5E5BC9C21000"

	Local $CodeBuffer = DllStructCreate("byte[" & BinaryLen($Opcode) & "]")
	DllStructSetData($CodeBuffer, 1, $Opcode)

	Local $Ouput = DllStructCreate("byte[" & BinaryLen($Data) & "]")
	Local $Ret = DllCall("user32.dll", "int", "CallWindowProc", "ptr", DllStructGetPtr($CodeBuffer), _
													"str", $Data, _
													"ptr", DllStructGetPtr($Ouput), _
													"int", 0, _
													"int", 0)

	Return BinaryMid(DllStructGetData($Ouput, 1), 1, $Ret[0])
EndFunc