<%

Public Function SendMail(sToName,sTo,sSubject,sBody)
	Dim Mailer
	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	Mailer.FromName   = "Classic ASP Framework"
	Mailer.FromAddress= "webmaster@claspdev.com"
	Mailer.AddRecipient sToName,sTo
	Mailer.RemoteHost = "mail.claspdev.com"
	Mailer.Subject    = sSubject
	Mailer.BodyText   = sBody
	If Mailer.SendMail Then
		SendMail = True
	Else
	  SendMail = False
	  Response.Write "--Mail send failure. Error was " & Mailer.Response
	End If
End Function

Public Function SendSupportMail(sFrom,sSubject,sBody)
	Dim Mailer
	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	Mailer.FromName   = "Classic ASP Framework"
	Mailer.FromAddress = sFrom
	Mailer.AddRecipient "WebMaster","webmaster@claspdev.com"
	Mailer.RemoteHost = "mail.claspdev.com"
	Mailer.Subject    = sSubject
	Mailer.BodyText   = sBody
	If Mailer.SendMail Then
		SendSupportMail= True
	Else
	  SendSupportMail = False
	  Response.Write "Mail send failure. Error was " & Mailer.Response
	End If
End Function

%>	