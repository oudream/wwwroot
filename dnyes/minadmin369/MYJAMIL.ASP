<%	
sub Jmail(email,topic,mailbody)
dim JMail
Set JMail=Server.CreateObject("JMail.Message")
JMail.Logging=True
JMail.Charset="gb2312"
JMail.ContentType = "text/html"
JMail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")
JMail.MailServerUserName = "postmaster@dnyes.com"
JMail.MailServerPassword = "winer031203"
JMail.From = "postmaster@dnyes.com" '发件人的信箱
JMail.FromName = "oudream"   '发件人的名字
JMail.Subject = topic  '邮件的主题
JMail.Body = mailbody   '邮件的内容
'==============================收件人的地址！
JMail.AddRecipient email         
JMail.AddRecipientBcc "oudream@21cn.com" '收件人的地址
JMail.Priority=3   '邮件级别1-5数字越大级别越高---3为普通邮件
JMail.Send("218.244.47.25") '红色变量是邮件服务器地址
JMail.Close
Set JMail=nothing 
end sub 
%>