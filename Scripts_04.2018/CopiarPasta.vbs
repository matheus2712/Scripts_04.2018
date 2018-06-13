On Error Resume Next

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set Shell = WScript.CreateObject("WScript.Shell")
Set ShellApplication = WScript.CreateObject("Shell.Application")


Set obj = ShellApplication.BrowseForFolder(0,"Selecione a pasta para ser copiada",0)

Dim strCaminho 


strPasta = obj.ParentFolder.ParseName(obj.Title).Path
 
WScript.Echo "Caminho da pasta a ser copiada " & strPasta

Set obj = ShellApplication.BrowseForFolder(0,"Informar em qual pasta deve ser feito backup",0)

strBackup = obj.ParentFolder.ParseName(obj.Title).Path


strLogFile = "logBackup.txt"
Set objLogFile = objFSO.OpenTextFile(strLogFile, 8, True, 0)

objLogFile.WriteLine "Backup realizado com sucesso" & now

If (objFSO.FolderExists(strPasta) = True) Then
   Set Folder = ObjFSO.GetFolder(strPasta)
		objLogFile.WriteLine   "origem: " & strPasta & " Backup: " & strBackup & " Data hora " & now
		ObjFSO.CopyFolder  strPasta, strBackup & "\", true
			

End if

Shell.Run(strBackup)
Shell.Run(strLogFile)

Set Shell = Nothing
Set ShellApplication = Nothing
Set ObjFSO = Nothing

wscript.quit