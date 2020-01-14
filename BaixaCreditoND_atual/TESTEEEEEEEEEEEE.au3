
;Array da base de dados
For $array = 0 To UBound($aArrayList) - 1
   For $i = 0 to $arrList.regions.lenght() -1
	  Local $region = $arrList.regions.Item($i)
	  For $j = 0 to Ubound($region) - 1
		 MsgBox(0,"","Teste")
	  Next
   Next
 Next
