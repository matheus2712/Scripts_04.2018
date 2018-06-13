set /p var2= Adicione o nome do computador
set /p var= Adicione o Caminho do arquivo
psexec %var2% -u matheus.camilo -p jwm@1995 -i -d -s %var%

pause

