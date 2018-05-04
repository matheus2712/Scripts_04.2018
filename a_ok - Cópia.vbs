Set Wshshell=CreateObject("Word.Basic")
'ativar a tecla prtScr
WshShell.sendkeys "{PRTSC}"
WshShell.sendkeys "(%{PRTSC})"

'time 1500 mile segundos
WScript.Sleep 150




set WshShell = CreateObject("WScript.Shell")
WshShell.Run "mspaint"
WScript.Sleep 2000
 

'Activating Paint Application
WshShell.AppActivate "untitled - Paint"
WScript.Sleep 1000
 
'Paste the captured Screenshot
WshShell.SendKeys "^v"
WScript.Sleep 500
 
'Save Screenshot
'WshShell.SendKeys "^b"
'WScript.Sleep 500
'WshShell.SendKeys "c:\test.bmp"
'WScript.Sleep 500
'WshShell.SendKeys "{ENTER}"