CD C:\Program Files\Microsoft Office\Office16

CD C:\Program Files (x86)\Microsoft Office\Office16

:menu

set /p var3= "Digite 1 para excluir outra ou 2 para Sair"

if %var3% == 1 goto :excluir
if %var3% == 2 goto :sair
goto end 
:excluir

cscript ospp.vbs /dstatus

set /p var2= "Adicione a chave"

cscript ospp.vbs /unpkey:%var2%

goto menu
:bas
echo "*****"
:end

