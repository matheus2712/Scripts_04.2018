Dim m

For m = 1 To 2 step 1 
 WScript.sleep 900


Set oShell = CreateObject("WScript.Shell")

sWallPaper = "\\server\shared\pp_atual.bmp"

' update in registry
oShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper", sWallPaper

oShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaperstyle", 3

' let the system know about the change
oShell.Run "%windir%\System32\RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters", 1, True



Next 