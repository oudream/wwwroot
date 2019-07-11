<html>
<head>
<title>
发送电子邮件
</title>
</head>
<body>
<%
towho=trim(request.form("towho")
subject=trim(request.form("subject")
content=trim(request.form("content")
if towho<>"" then
    set myMail=server.createObject("CDONTS.Newmail")
    myMail.from="fang@abc.com"
    myMail.to=towho
    myMail.subject=subject
    myMail.body=content
    myMail.send
    set myMail=nothing 
    Response.Write "你的邮件已经发往" & towho
else
    Response.Write "你请输入你所要发往的邮箱地址"
end if
</body>
</html>