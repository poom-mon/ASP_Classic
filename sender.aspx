<%@ Page ValidateRequest="false"%>
<%@ Import Namespace="System.Net.Mail" %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<title></title>
<%
	Dim Sender, SenderName, Recipient, RecipientCC, RecipientBCC, Subject, Body, Attach ,MailServerUserName, MailServerPassword, Debug, Redirect, Result as string
	Dim arrRecipient() as string 
	Dim arrRecipientCC() as string 
	Dim arrRecipientBCC() as string 
	Dim arrAttach() as string 
	Dim i as integer
	Sender = request("Sender")
	SenderName = request("SenderName")
	if SenderName <> "" then
		Sender = SenderName &" <"& Sender &">"
	end if
	Recipient = request("Recipient")
	arrRecipient = split(Recipient, ",")
	RecipientCC = request("RecipientCC")
	arrRecipientCC = split(RecipientCC, ",")
	RecipientBCC = request("RecipientBCC")
	arrRecipientBCC = split(RecipientBCC, ",")
	Subject = request("Subject")
	Body = request("Body")
	Body = replace(Body, ".[.", "<")
	Body = replace(Body, ".].", ">")		
	Attach = request("Attach")
	Attach = replace(Attach, " ", "")		
	arrAttach = split(Attach, ",")
	MailServerUserName = request("MailServerUserName")
	MailServerPassword = request("MailServerPassword")
	if MailServerUserName = "" then
		MailServerUserName = "internal@silkspan.com"
		MailServerPassword = "internalss"
	end If
	Debug = request("Debug")
	Redirect = request("Redirect")
	
	Dim objEmail as New MailMessage()
	Dim Attachment As System.Net.Mail.Attachment
   
	objEmail.IsBodyHtml = True
	objEmail.Priority = MailPriority.Normal
	objEmail.From = New MailAddress(Sender)
	for i = 0 to ubound(arrRecipient)
		if arrRecipient(i) <> "" then
			objEmail.To.Add(New MailAddress(arrRecipient(i)))
		end if
	next

	for i = 0 to ubound(arrRecipientCC)
		if arrRecipientCC(i) <> "" then
			objEmail.CC.Add(New MailAddress(arrRecipientCC(i)))
		end if
	next

	for i = 0 to ubound(arrRecipientBCC)
		if arrRecipientBCC(i) <> "" then
			objEmail.BCC.Add(New MailAddress(arrRecipientBCC(i)))
		end if
	next

	objEmail.Subject = Subject

	for i = 0 to ubound(arrAttach)
		if arrAttach(i) <> "" Then
			Attachment = New System.Net.Mail.Attachment(server.mappath(arrAttach(i)))

			Attachment.ContentId = i + 1
			Body = replace(Body, "<img src='" + arrAttach(i) +"'", "<img src='cid:" + Attachment.ContentId +"'")
            Body = Replace(Body, " background='" + arrAttach(i) + "'", " background='cid:" + Attachment.ContentId + "'")
            Body = Replace(Body, "jobsnews", "สื่อสิ่งพิมพ์")
			objEmail.Attachments.Add(Attachment)				
		end if
	next

	objEmail.Body = Body

	Dim SmtpMail As New SmtpClient()
    SmtpMail.Credentials = New System.Net.NetworkCredential(MailServerUserName, MailServerPassword)
    SmtpMail.Port = 587
	SmtpMail.EnableSsl = True

	If Debug = "T" Then
		SmtpMail.Send(objEmail)
	else
		Try
			SmtpMail.Send(objEmail)
			response.write("<font color='#006600' style='font-size: 8pt;'>Send Mail Complete</font>")
			Result = "Complete"
		Catch e As Exception
			response.write("<font color='#CC0000' style='font-size: 8pt;'>")
			response.write("Send Mail Incomplete<br>"& e.Message)
			response.write("<br>Sender : "& Sender)
			response.write("<br>MailServerUserName : "& MailServerUserName)
			response.write("</font>")
			Result = "Incomplete"
			for each a As Attachment in objEmail.Attachments
				a.Dispose()
			next
		End Try
	End If

	SmtpMail = nothing

	for each a As Attachment in objEmail.Attachments
		a.Dispose()
	next
	objEmail.Attachments.Dispose()
	objEmail.Dispose()
	objEmail = Nothing

	If Redirect <> "" and Result = "Complete" Then
		response.write("<form name='Fm' method='post' action='"& Redirect &"'>")
		response.write("<input type='hidden' name='result' value='"& Result &"'>")
		response.write("</form>")

		response.write("<script type='text/javascript'>")
		response.write("document.Fm.submit();")
		response.write("</script>")
	End if
%>
</body>
</html>