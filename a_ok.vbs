Set Wshshell=CreateObject("Word.Basic")
'ativar a tecla prtScr
WshShell.sendkeys "{PRTSC}"

WshShell.sendkeys "(%{PRTSC})"
'time 1500 mile segundos
WScript.Sleep 150



'Sys.Desktop.Picture.SaveToFile "C:\screenshot.png"

'WshShell.sendKeys.Send("^v")


set WshShell = CreateObject("WScript.Shell")
WshShell.AppActivate "Untitled - Paint" 
WScript.Sleep 150 

WshShell.sendkeys "^(v)" 
WScript.Sleep 150 

WshShell.sendkeys "^ (s)" 
WScript.Sleep 150

WshShell.sendkeys "testing.jpg" 
WScript.Sleep 150 

WshShell.sendkeys "% (s)" 
WScript.Sleep 150 
