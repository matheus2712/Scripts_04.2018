
set /p var2= "Adicione o nome do computador ou usuario  "
find /i "%var2%" c:/lst.txt

set /p var1= Adicione o IP :
set var3=\\192.168.0.%var1%

psexec %var3% -u matheus.camilo -p ma*646921640 -i -d -s \\server\shared\Skype.exe

pause