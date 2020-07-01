'!!!!! USE ON YOUR OWN RISK !!!!!
'=== Send SMTP Email ==================================
' CREATOR: DJIA | eniacjia@gmail.com
' BUILD VERSION: [2020-07-01]
' FUNCTION:  Send SMTP Email with  up to 2 attachments
' PARAMETER LIST -------------------------------------
'   EmailFrom :  Email sent from
'   EmailTo:  Email to, email1;email2 format 
'   EmailCC: Email CC: Email to, email1;email2 format
'   Subject: Subject of the email
'   EmailBody:  Content of the email
'   SmtpServer:  SMTP Server IP or host name
'   FileAttachment:  empty or path of the file
'   FileAttachment2: empty or path of the file
' UPDATE LOG ------------------------------------------ 
'  [2020-07-01] DJ: Update the descriptions
'======================================================
Sub SendEmail(EmailFrom, EmailTo, EmailCC, Subject, EmailBody, SmtpServer, FileAttachment, FileAttachment2)
    On Error Resume Next
    Dim MyEmail
    Set MyEmail = CreateObject("CDO.Message")
    MyEmail.From = EmailFrom
    MyEmail.To = EmailTo
    MyEmail.Cc = EmailCC
    MyEmail.Subject = Subject
    MyEmail.Textbody = EmailBody
    MyEmail.AddAttachment FileAttachment
	If FileAttachment2<>"" then
		MyEmail.AddAttachment FileAttachment2
	End If
    MyEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    MyEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SmtpServer
    MyEmail.Configuration.Fields.Update
    MyEmail.Send
    Set MyEmail = Nothing
End Sub
