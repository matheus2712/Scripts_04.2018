Set Wshshell=CreateObject("Wscript.shell")
'ativar a tecla prtScr
WshShell.sendkeys "{PRTSC}"

WshShell.sendkeys "(%{PRTSC})"
'time 1500 mile segundos
WScript.Sleep 150
'Sys.Desktop.Picture.SaveToFile "C:\screenshot.png"




tes

 
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
