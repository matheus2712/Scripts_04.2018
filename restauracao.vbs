rp="Ponto de restauto criado pelo atalho " & WScript.ScriptName
GetObject("winmgmts:\\.\root\default:Systemrestore").CreateRestorePoint rp, 0, 100
Msgbox("Ponto de restauro criado!")