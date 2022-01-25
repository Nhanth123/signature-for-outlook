On Error Resume Next


Set objNetwork = CreateObject("Wscript.Network")

StartNode = "" 'replace it with your DC for example: DC=rochdalespears,DC=com
strAccount = objNetwork.UserName

Set objCommand = CreateObject("ADODB.Command")
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
objCommand.ActiveConnection = objConnection
SearchScope = "subtree"

FilterString = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=" & strAccount & "))"
Attributes = "adspath"

LDAPQuery = "<LDAP://" & StartNode & ">;" & FilterString & ";" _
        & Attributes & ";" & SearchScope

objCommand.CommandText = LDAPQuery
objCommand.Properties("Page Size") = 1500
objCommand.Properties("Timeout") = 30
objCommand.Properties("Cache Results") = False

Set objRecordset = objCommand.Execute

If Not objRecordset.EOF Then
   objRecordset.MoveFirst

   Do Until objRecordset.EOF
      strUserPath = objRecordset.Fields("ADsPath").Value
      Set objUser = GetObject(strUserPath)
      
	  displayName = objUser.DisplayName
	  phoneNumber = objUser.TelephoneNumber
	  email = objUser.mail
     
      objRecordset.MoveNext
   Loop
End If

objRecordset.Close
objConnection.Close

fullName = Split(displayName, "/")(0)
title = Trim(Split(displayName, "/")(1))
mailAddress = "mailto:"& email


UrlRSGPanel= "https://res.cloudinary.com/dwyjr7wio/image/upload/v1635563624/RSG/RSGPanel_xwzntv.jpg"
UrlRSGLinkedin = "https://res.cloudinary.com/dwyjr7wio/image/upload/v1635573249/RSG/RSG_LinkedinIcon_rht178.png"
UrlRSGMail = "https://res.cloudinary.com/dwyjr7wio/image/upload/v1635572713/RSG/RSG_MailIcon_hwzkyg.png"
UrlRSGHomepage = "https://res.cloudinary.com/dwyjr7wio/image/upload/v1635573293/RSG/RSG_HomepageIcon_r1kota.png"
UrlRochdale = "https://www.rochdalespears.com/"
UrlSonder = "https://sonderliving.com/"
UrlLinkedin = "https://www.linkedin.com/company/rochdale-spears-co.-ltd/mycompany/"
Set objWord = CreateObject("Word.Application")

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
objSelection.Font.Name = "Aharoni"
objSelection.Font.Size = 11
objSelection.Font.Bold = True


Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature

Set objSignatureEntries = objSignatureObject.EmailSignatureEntries



' BOF signature
objSelection.Range(0, 0).Select
objSelection.Font.Name = "Aharoni"
objSelection.Font.Size = 11
objSelection.Font.Bold = True
objSelection.TypeText ""

objSelection.TypeParagraph()
objSelection.Font.Name = "Avenir Next Medium"
objSelection.Font.Size = 11
objSelection.Font.Bold = False
objSelection.TypeText fullName
objSelection.TypeParagraph()

objSelection.Font.Name = "Avenir Next Ultra Light"
objSelection.Font.Size = 11
objSelection.TypeText title
objSelection.TypeParagraph()

objSelection.Font.Size = 10
objSelection.TypeText ""& phoneNumber
objSelection.TypeParagraph()

objDoc.Tables.Add objRange, 1,4
Set objTable = objDoc.Tables(1)
objTable.Cell(1,1).select
Set objShapeLinkedin = objSelection.InLineShapes.AddPicture(UrlRSGLinkedin)
objDoc.Hyperlinks.Add objShapeLinkedin, UrlLinkedin

objTable.Cell(2,1).select
objSelection.TypeText "		"

objTable.Cell(3,1).select
Set objShapeRSG = objSelection.InLineShapes.AddPicture(UrlRSGHomepage)
objDoc.Hyperlinks.Add objShapeRSG, UrlRochdale

objTable.Cell(4,1).select
objSelection.TypeText "		"

Set objShapeMail = objSelection.InLineShapes.AddPicture(UrlRSGMail)
objDoc.Hyperlinks.Add objShapeMail, mailAddress


objSelection.TypeParagraph()
Set objShape = objSelection.InLineShapes.AddPicture(UrlRSGPanel)
objDoc.Hyperlinks.Add objShape, UrlRochdale
objSelection.Font.Name = "Avenir Next Medium"
objSelection.Font.Size = 10
objSelection.Font.Bold = False
objSelection.TypeParagraph()
objSelection.Font.Color = RGB(0,0,0)

objLink = objSelection.Hyperlinks.Add(objSelection.Range,UrlSonder,,, "SONDER LIVING")
objSelection.TypeText " a member of the Rochdale Spears Group."
objSelection.TypeParagraph()
objSelection.Font.Name = "Avenir Next Ultra Light"
objSelection.Font.Size = 8
objSelection.TypeText "CONFIDENTIALITY NOTICE: This email, including any attachments, is confidential. If you are not the intended recipient please notify the sender immediately, and please delete it and any attachments from your system without reading the content: you should not copy it or use it for any purpose or disclose its contents to any other person. There is no intent on the part of the sender to waive any privilege that may attach to this communication. Thank you for your cooperation."

objSelection.TypeParagraph()



' EOF signature

Set objSelection = objDoc.Range()

objSignatureEntries.Add "MyCompanySignature", objSelection
objSignatureObject.NewMessageSignature = "MyCompanySignature"
objSignatureObject.ReplyMessageSignature = "MyCompanySignature"

objDoc.Saved = True

'Wscript.Echo "Success!"&chr(13)&chr(10)&"Simply restart Outlook and enjoy your new signature."

objWord.Quit