<%
	sub Jmail(email,topic,mailbody)
	on error resume next
	dim JMail
	Set JMail=Server.CreateObject("JMail.SMTPMail")
	JMail.Logging=True
	JMail.Charset="gb2312"
	JMail.ContentType = "text/html"
	JMail.ServerAddress=Forum_info(4)
	JMail.Sender=Forum_info(5)
	JMail.Subject=topic
	JMail.Body=mailbody
	JMail.AddRecipient email
	JMail.Priority=1
	'JMail.MailServerUserName = "shatan@dvbbs.net" '您的邮件服务器登录名
	'JMail.MailServerPassword = "nihaoma" '登录密码
	JMail.Execute 
	Set JMail=nothing 
	if err then 
	SendMail=err.description
	err.clear
	else
	SendMail="OK"
	end if
	end sub

	sub Cdonts(email,topic,mailbody)
	on error resume next
	dim  objCDOMail
	Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
	objCDOMail.From =Forum_info(5)
	objCDOMail.To =email
	objCDOMail.Subject =topic
	objCDOMail.BodyFormat = 0 
	objCDOMail.MailFormat = 0 
	objCDOMail.Body =mailbody
	objCDOMail.Send
	Set objCDOMail = Nothing
	if err then 
	SendMail=err.description
	err.clear
	else
	SendMail="OK"
	end if
	end sub

	sub aspemail(email,topic,mailbody)
	on error resume next
	dim mailer,recipient,sender,subject,message
	dim mailserver,result
	Set mailer=Server.CreateObject("ASPMAIL.ASPMailCtrl.1")  
	recipient=email
	sender=Forum_info(5)
	subject=topic
	message=mailbody
	mailserver=Forum_info(4)
	result=mailer.SendMail(mailserver, recipient, sender, subject, message)
	if err then 
	SendMail=err.description
	err.clear
	else
	SendMail="OK"
	end if
	end sub

	sub Emailbody(body)
	Emailbody=Emailbody &"<style>A:visited {	TEXT-DECORATION: none	}"
	Emailbody=Emailbody &"A:active  {	TEXT-DECORATION: none	}"
	Emailbody=Emailbody &"A:hover   {	TEXT-DECORATION: underline overline	}"
	Emailbody=Emailbody &"A:link 	  {	text-decoration: none;}"
	Emailbody=Emailbody &"A:visited {	text-decoration: none;}"
	Emailbody=Emailbody &"A:active  {	TEXT-DECORATION: none;}"
	Emailbody=Emailbody &"A:hover   {	TEXT-DECORATION: underline overline}"
	Emailbody=Emailbody &"BODY   {	FONT-FAMILY: 宋体; FONT-SIZE: 9pt;}"
	Emailbody=Emailbody &"TD	   {	FONT-FAMILY: 宋体; FONT-SIZE: 9pt	}</style>"
	Emailbody=Emailbody &"<TABLE border=0 width='95%' align=center><TBODY><TR>"
	Emailbody=Emailbody &"<TD valign=middle align=top>"
	Emailbody=Emailbody & body
	Emailbody=Emailbody &"</TD></TR></TBODY></TABLE><br><hr width=95% size=1>"
	Emailbody=Emailbody & Copyright & " &nbsp;&nbsp; " & Version
	Emailbody=mailbody
	end sub
%>