Set FSo = CreateObject("Scripting.FileSystemObject")
set u = CreateObject("WScript.Network")
getUser = u.username
Set oLog = FSO.OpenTextFile("\\server\shared\log\log.txt", 8, True)
sInicio = Now()


oLog.WriteLine "**** Inicio do Log: " & sInicio & "*********"	
oLog.WriteLine " > " & getUser
oLog.WriteLine "**** Finalizado Log ********************"
oLog.WriteLine ""
oLog.Close