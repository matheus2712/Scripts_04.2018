Dim j

For j = 1 To 5 step 1 
 
WScript.sleep 900

Set oShell = CreateObject("WScript.Shell")

sWallPaper = "\\server\shared\pp_atual.bmp"

' update in registry
oShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper", sWallPaper

oShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaperstyle", 0

' let the system know about the change
oShell.Run "%windir%\System32\RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters", 2, True



Next 

Dim k 

For k = 1 To 5 step 1 
 WScript.sleep 900


Set oShell = CreateObject("WScript.Shell")

sWallPaper = "\\server\shared\pp_atual.bmp"

' update in registry
oShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper", sWallPaper

oShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaperstyle", 1

' let the system know about the change
oShell.Run "%windir%\System32\RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters", 2, True



Next 

Dim l

For l = 1 To 5 step 1 
 WScript.sleep 900


Set oShell = CreateObject("WScript.Shell")

sWallPaper = "\\server\shared\pp_atual.bmp"

' update in registry
oShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper", sWallPaper

oShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaperstyle", 4

' let the system know about the change
oShell.Run "%windir%\System32\RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters", 2, True



Next 

Dim m

For m = 1 To 5 step 1 
 WScript.sleep 900


Set oShell = CreateObject("WScript.Shell")

sWallPaper = "\\server\shared\pp_atual.bmp"

' update in registry
oShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper", sWallPaper

oShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaperstyle", 3

' let the system know about the change
oShell.Run "%windir%\System32\RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters", 1, True



Next 