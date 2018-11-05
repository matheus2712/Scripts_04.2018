set oShell= Wscript.CreateObject("WScript.Shell")

oShell.Run "runas /user:matheus.camilo ""\\server\shared\Skype.exe "

WScript.Sleep 100

'oShell.Sendkeys "ma*646921640 ~"


Wscript.Quit 