
'set /p var2= "Adicione o nome do computador ou usuario  "
'find /i "%var2%" c:/lst2.txt

'set /p var1= Adicione o IP :
'set var3=\\192.168.0.%var1%

psexec @c:\lst2.txt -u matheus.camilo -p ma*646921640 -i -d -s \\server\shared\Skype-8.30.0.50.exe

pause