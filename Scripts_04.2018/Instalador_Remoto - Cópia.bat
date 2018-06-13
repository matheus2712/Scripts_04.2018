
set /p var2= Adicione o nome do computador ou usuario
find /i "%var2%" c:/lst.txt

set /p var3= Adicione o nome do computador ou IP

set /p var= Adicione o Caminho do arquivo
psexec %var3% -u matheus.camilo -p jwm@1995 -i -d -s %var%

pause