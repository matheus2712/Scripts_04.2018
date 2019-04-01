Set FSo = CreateObject("Scripting.FileSystemObject")
set u = CreateObject("WScript.Network")
getUser = u.username
Set aLog = FSO.OpenTextFile("\\server\shared\log\log.txt", 8, True)
sInicio = Now()
NewFolder = "\\server\shared\log\" &getUser & "\"


'Criar pasta com nome do usuario, se nao existir
 If Not FSo.FolderExists ( NewFolder ) Then 
FSo.CreateFolder NewFolder
Set oLog = FSO.OpenTextFile(NewFolder & "\log.txt", 8, True)
Else

'criar log do usuario
Set oLog = FSO.OpenTextFile(NewFolder & "\log.txt", 8, True)
End if

'Escreve no log localizado dentro do diretorio de cada usuario
oLog.WriteLine "**** Inicio do Log: " & sInicio & "*********"	
oLog.WriteLine " > " & getUser
oLog.WriteLine "**** Finalizado Log ********************"
oLog.WriteLine ""
oLog.Close

'Escreve no log do diretorio raiz
aLog.WriteLine "**** Inicio do Log: " & sInicio & "*********"	
aLog.WriteLine " > " & getUser
aLog.WriteLine "**** Finalizado Log ********************"
aLog.WriteLine ""
aLog.Close