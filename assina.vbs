'--------------------------------------------------
' Script de Configuração das Assinaturas do OUTLOOK
'--------------------------------------------------
On Error Resume Next

Set objSysInfo = CreateObject("ADSystemInfo")
strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)

'----------------------------------------------------------------------------' Campos selecionados no AD
' DisplayName, Title, HomePhone, FacsimileTelephoneNumber, IpPhone, Mobile e Pager
'----------------------------------------------------------------------------strNome = objUser.DisplayName
strDpto = objUser.Title
strTelDireto = objuser.homePhone
strCelular = objuser.mobile
strNextel = objuser.facsimileTelephoneNumber
strIdNextel = objuser.pager
strMSN = objuser.ipPhone
strEMail = objuser.mail
strLogo = "\\server\shared\Assinatura novas\ester.silva2018-01.jpg"

'----------------------
' Criação de Assinatura
'----------------------
Set objWord = CreateObject("Word.Application")
objWord.Visible = False

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
Set objRange = objDoc.Range()

Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

'---------------------------------------------------------------------------------
' Define fonte e corpo Email
'---------------------------------------------------------------------------------
'objSelection.Font.Size = "10"
'objSelection.Font.Name = "Arial"
'objSelection.Font.Bold = False
'objSelection.Font.Color = RGB(0,70,125)
'objSelection.TypeParagraph()
'objSelection.TypeParagraph()

'---------------------------------------------------------------------------------
' Define fonte Nome
'---------------------------------------------------------------------------------
objSelection.Font.Size = "12"
objSelection.Font.Name = "Arial"
objSelection.Font.Bold = True
objSelection.Font.Color = RGB(0,70,125)
objSelection.TypeText " " & strNome
objSelection.TypeParagraph()

'---------------------------------------------------------------------------------
' Define fonte Cargo
'---------------------------------------------------------------------------------
objSelection.Font.Size = "10"
objSelection.Font.Name = "Arial"
objSelection.Font.Bold = True
objSelection.Font.Color = RGB(102,102,102)
objSelection.TypeText " " & strDpto
objSelection.TypeParagraph()

'---------------------------------------------------------------------------------
' Define fonte demais informações
'---------------------------------------------------------------------------------
objSelection.Font.Size = "10"
objSelection.Font.Name = "Arial"
objSelection.Font.Bold = False
objSelection.Font.Color = RGB(102,102,102)

'---------------------------------------------------------------------------------
' Condição para inserir telefones SE existir
'---------------------------------------------------------------------------------
' Condição para inserir Telefono Direto
	if strTelDireto <> "" Then
		objSelection.TypeText " Telefone Direto: " & strTelDireto
		objSelection.TypeParagraph()
	End If
' Condição para inserir Celular
	If strCelular <> "" Then
		objSelection.TypeText " Telefone Celular: " & strCelular
		objSelection.TypeParagraph()
	End If

' Condição para inserir Nextel
	If strNextel <> "" Then
		objSelection.TypeText " Nextel: " & strNextel
		objSelection.TypeParagraph()
	End If

' Condição para inserir ID Nextel
	If strIdNextel <> "" Then
		objSelection.TypeText " ID Nextel: " & strIdNextel
		objSelection.TypeParagraph()
	End If

' Condição para inserir MSN
	If strMSN <> "" Then
		objSelection.TypeText " IM MSN: " & strMSN
		objSelection.TypeParagraph()
	End If

' Cria link e exibe email
objSelection.TypeText " ":
objSelection.Hyperlinks.Add objSelection.range, "mailto:" & strEMail, , , strEMail
objSelection.TypeParagraph()
'---------------------------------------------------------------------------------


'---------------------------------------------------------------------------------
' Inseri JPEG
'---------------------------------------------------------------------------------
objSelection.InlineShapes.AddPicture(strLogo) 


objSignatureEntries.Add "Signature", objRange
objSignatureObject.NewMessageSignature = "Signature"
objSignatureObject.ReplyMessageSignature = "Signature"

objDoc.Saved = True
objWord.Quit