REM En cas général les "#" servent à faire des commentaires comme ici
echo Lancement scrap de PronoSoft
PATH=./Import/xlrd-0.7.1:./Import/xlwt-0.7.2:./Import/pyexcelerator-0.6.4.1:%PATH%
python ./src/PS.py %1 %2 %3 %4 %5
 
