
#include "OO_JSON.au3"
#include <Array.au3>
#include <Excel.au3>

Local $sFullFile = "C:\Users\wesley.valadao\Desktop\ContasJson.txt"
Local $file = FileOpen($sFullFile)
Local $Text = FileRead($file)
Local $oJSON = _OO_JSON_Init()

$jsObj = $oJSON.strToObject($Text) ; JSON encode key value

Global $iQtdCaracterConta

;~ _mainExcel()

_Json()

Func _Json()

Global $i
Global $a

   ;~ ConsoleWrite($jsObj.regions.length())
   For $i = 0 to $jsObj.regions.length() -1
   ;~    For $a = 0 TO $jsObj.regions.item($i).lines.item(1).words.item(0).text & @CRLF)
	  For $a = 0 TO $jsObj.regions.item($i).lines.length()-1
		 _Valida_Conta($jsObj.regions.item($i).lines.item($a).words.item(0).text)
	  Next
   Next

EndFunc

Func _Valida_Conta($ContaAtlys)

ConsoleWrite($ContaAtlys)

If StringIsInt($ContaAtlys) = True Then
   MsgBox(0,"","Só numeros")
   _ArrayInsert($ContaAtlys,0,$ContaAtlys)
Else
   $Letra = StringMid($ContaAtlys,1,10)

EndIf

_ArrayDisplay($ContaAtlys)

EndFunc


Func _CaracteresEspeciais($caract)
	local $codiA = "àáâãäèéêëìíîïòóôõöùúûüÀÁÂÃÄÈÉÊËÌÍÎÒÓÔÕÖÙÚÛÜçÇñÑ"
	local $codiB = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN"
	local $temp = $caract
	local $result = ""

		For $i = 1 To StringLen($temp)
			$p = StringInStr($codiA, StringMid($temp, $i, 1))
			If $p > 0 Then
				$result = $result & StringMid($codiB, $p, 1)
			else
				$result = $result & StringMid($temp, $i, 1)
			endif
		Next

	Return $result
 EndFunc

Func _mainExcel()

   ;INSTANCIAR OBJETO
   Global $oExcelObj = ObjCreate("Excel.Application")
   $oExcelObj.visible = False

   ;ABRIR WORKBOOK
   Global $oWbk =  $oExcelObj.WorkBooks.Open(@ScriptDir & "\Relatório Robo.xlsx")

   ;CARREGAR WORKSHEET CADASTRO - A planilha deve ser a primeira da esquerda
   Global $sht =  $oWbk.Worksheets(1)

   ;QUANTIDADE DE LINHAS
   Global $iRowCount = $sht.UsedRange.Rows.Count

   ;QUANTIDADE COLUNAS
   Global $iColumnCount = $sht.UsedRange.Columns.Count

   ;POSICIONAR APOS A ULTIMA COLUNA PARA STATUS
;~	Global $iLastColumn = _Excel_ColumnToLetter($iColumnCount+1)

   ;RANGE
   Global $iRange = "A2" & ":H" & $iRowCount
   ;Global $iRange = "A2" & ":BC" & $iRowCount

   ;ARRAY DO RANGE
   Global $aArrayList = _Excel_RangeRead($oWbk,$sht,$iRange,1,Default)

;~    _ArrayDisplay($aArrayList)

   ;_ArrayDisplay($aArrayList)

   ;Valida se a lista é um Array e se possuí dados
   If Not IsArray($aArrayList) Or Ubound($aArrayList) = 0 then
	  MsgBox($MB_ICONERROR,$sApp[$eAutomatioName],"Não foram identificados registros no arquivo selecionado!")
	  _Exit()
   EndIf

EndFunc

;Fecha a planilha e encerra a aplicação Excel
;============================================
Func _PLANILHA_FECHAR()
	$oWbk.close
	$oExcelObj.quit
	$oExcelObj = Null
 EndFunc

;~ $json_text = _OO_JSON_Read_File($Text)
;~ ConsoleWrite($jsObj)

;~ 6400010650
;~ 3155010B10
;~ 6400011240
;~ 6400011245
;~ 6400011235