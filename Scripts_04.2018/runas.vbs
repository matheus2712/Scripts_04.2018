set oShell= Wscript.CreateObject("WScript.Shell")

oShell.Run "runas /user:matheus.camilo ""C:\Program Files (x86)\Microsoft\Skype for Desktop\Skype.exe "

WScript.Sleep 100

'oShell.Sendkeys "jwm@1995 ~"


Wscript.Quit 