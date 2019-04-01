Option explicit

Dim oShell


set oShell= Wscript.CreateObject("WScript.Shell")

If Day(Date) = "28" or Day(Date) = "25" Then


GetObject("winmgmts:\\.\root\default:Systemrestore").CreateRestorePoint "test", 0, 100
Wscript.Echo "feito"

Else
    Wscript.Echo "nao feito"
End if

Wscript.Quit 