dir /s/b *.txt > log.txt
winscp.com /script=WinSCP_UpLoad.cfg
mv *.txt /"C:\ExtracaoBRNotas_PDF"
move /Y C:\ExtracaoBRNotas_PDF\NotasPDF\*.csv C:\ExtracaoBRNotas_PDF\Processado\
del *.txt
#pause
exit