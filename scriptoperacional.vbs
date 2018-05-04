Set objShell = WScript.CreateObject("WScript.Shell")
Set objEnvironment = objShell.Environment("Process")
'strWinDir=objEnvironment.Item("WinDir")
strHomePath=objEnvironment.Item("HomePath")
strHomeDrive=objEnvironment.Item("HomeDrive")
Set objNet = CreateObject("Wscript.Network") 

'(Sincroniza Data/Hora Servidor)
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "net time \\server /set /y", 0, TRUE

'Mapeandro Impressora Operacional Duplex
Set objimpressora = CreateObject("wscript.Network")
objimpressora.AddWindowsPrinterConnection "\\server\Operacional_Duplex"
objimpressora.SetDefaultPrinter "\\server\operacional_Duplex"

'Atalho ESL
Set lnk = objShell.CreateShortcut(strHomeDrive & strHomePath & "\Desktop\Sistema ESL.lnk") 
lnk.TargetPath = strWinDir & "\\srvbd\Transportes\principal\TelaEntrada6.exe"
lnk.Description = "Sistema ESL"
lnk.IconLocation = strWinDir & "\\srvbd\Transportes\principal\TelaEntrada6.exe"
lnk.WindowStyle = "2"
lnk.WorkingDirectory = strWinDir & "\\srvbd\Transportes\principal"
lnk.HotKey = "CTRL+SHIFT+C"
lnk.Save

'Atalho Sistema Qualidade
Set lnk = objShell.CreateShortcut(strHomeDrive & strHomePath & "\Desktop\RNCI.lnk") 
lnk.TargetPath = strWinDir & "\\server\JWM_ARQUIVOS\SIG JWM\Documentos_Normativos\Registros\Sistema Qualidade_2014\SIG-Controle da Qualidade.accdb"
lnk.Description = "Sistema Qualidade"
lnk.IconLocation = strWinDir & "\\server\JWM_ARQUIVOS\SIG JWM\Documentos_Normativos\Registros\Sistema Qualidade_2014\sgq.ico"
lnk.WindowStyle = "2"
lnk.WorkingDirectory = strWinDir & "\\server\JWM_ARQUIVOS\SIG JWM\Documentos_Normativos\Registros\Sistema Qualidade_2014"
lnk.HotKey = "CTRL+SHIFT+R"
lnk.Save

'Atalho Intranet Qualidade
Set lnk = objShell.CreateShortcut(strHomeDrive & strHomePath & "\Desktop\Intranet SIG.lnk") 
lnk.TargetPath = strWinDir & "\\server\JWM_ARQUIVOS\SIG JWM\Documentos_Normativos\Registros\Intranet_SIG_JWM\index.htm"
lnk.Description = "Intranet SIG"
lnk.IconLocation = strWinDir & "\\server\JWM_ARQUIVOS\SIG JWM\Documentos_Normativos\Registros\Intranet_SIG_JWM\intranet.ico"
lnk.WindowStyle = "2"
lnk.WorkingDirectory = strWinDir & "\\server\JWM_ARQUIVOS\SIG JWM\Documentos_Normativos\Registros\Intranet_SIG_JWM\index.htm"
lnk.HotKey = "CTRL+SHIFT+Q"
lnk.Save

set WshShell = WScript.CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop")
set oUrlLink = WshShell.CreateShortcut(strDesktop & "\GLPI.url")
oUrlLink.TargetPath = "http://192.168.0.71/glpi/index.php?noAUTO=1"
oUrlLink.Save

Set lnk = objShell.CreateShortcut(strHomeDrive & strHomePath & "\Desktop\NOVO CHAMADO.lnk") 
lnk.TargetPath = strWinDir & "mailto:suporte@jwmtransportes.com.br"
lnk.Description = "Criar Chamado"
lnk.IconLocation = strWinDir & "\\server\shared\glpi.ico"
lnk.WindowStyle = "2"
lnk.WorkingDirectory = strWinDir & ""
lnk.HotKey = "CTRL+SHIFT+C"
lnk.Save

'Mapeando Unidade de Rede
set net = createobject("wscript.network")
Set FSODrive= CreateObject("Scripting.FileSystemObject")
If not FSODrive.DriveExists("J:") Then
Set NW = CreateObject("WScript.Network")
NW.MapNetworkDrive "J:", "\\server\JWM_ARQUIVOS", False
End If


'Aplica o PROXY e adiciona URLs do Office
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "\\server\SYSVOL\jwm.local\script\atualiza_proxy_1.vbs"

'Recebe 
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "\\server\SYSVOL\jwm.local\script\Get_Login.vbs"



'papel de parede
Set oShell = CreateObject("WScript.Shell")
sWallPaper = "\\server\shared\pp_atual.bmp"
' update in registry
oShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper", sWallPaper
' let the system know about the change
oShell.Run "%windir%\System32\RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters", 2, True



'Mapeando Unidade de Rede
set net = createobject("wscript.network")
Set FSODrive= CreateObject("Scripting.FileSystemObject")
If not FSODrive.DriveExists("S:") Then
Set NW = CreateObject("WScript.Network")
NW.MapNetworkDrive "S:", "\\Server\Scanner_Operacional", False
End If
wscript.quit

