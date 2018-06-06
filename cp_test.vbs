Dim oFS
Dim sDstDir
Dim sSrcDir

Set oFS = CreateObject("Scripting.FileSystemObject")


sDstDir = "C:\Skype"

'Arquivos que serão copiados
sSrcDir = "\\server\shared\Skype" 



' se nao existir, ira criar

If Not oFS.FolderExists ( sDstDir) Then 
'Cria e Copia

oFS.CreateFolder sDstDir
oFS.CopyFolder sSrcDir, sDstDir, TRUE
 Else
'nao faz nada
sDstDir = "C:\Skype"

End If
