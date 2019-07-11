<%@language=vbscript%>
<html>
<head>
</head>
<body bgcolor=#c1f7d8>
<center>
<%
dim strname,strpasswod1,strpassword2,strnote,stremail,strhomepage,strpassword
strname=request.form("Name")
strpassword1=request.form("password1")
strpassword2=request.form("password2")
stremail=request.form("email")
strhomepage=request.form("homepage")
strnote=request.form("note")
if strname="" then
   response.write "帐号不能为空<p></p>"
%>
<a href=javascript:history.back()>上一页</a>
<%
   response.end
end if
if strpassword1="" then
   response.write "密码不能为空<p></p>"
%>
<a href=javascript:history.bakc()>上一页</a>
<%
   Response.end
end if
if strpassword1=strpassword2 then
    strpassword=strpassword1
else
    response.write "确认两次密码一致<br><br>"
%>
<a href=javascript:history.back()>上一页</a>
<%
response.end
end if
strsql="select * from user where username='" & strname & "'"
strdsn="dsn=bbs;uid=feng;pwd=feng"
set rs=server.createobject("adodb.recordset")
rs.open strsql,strdsn,1,3
if not (rs.eof and rs.eof) then
    response.write "你所用的帐号已经存在，请发变帐号<br><br>"
%>
<a href=javascript:history.back()>上一页</a>
<%
  response.end
end if
rs.close
set rs=nothing
strtable="user"
strdsn="dsn=bbs;uid=feng;pwd=feng"
set rs=server.createobject("adodb.recordset")
rs.open strtable,strdsn,1,3
rs.addnew
rs("username")=strname
rs("userpassword")=strpassword1
rs("useremail")=stremail
rs("userhomepage")=strhomepage
rs("usernote")=strnote
rs.update
rs.close
set rs=nothing
response.write "祝贺你，你已经申请成功"
%>
<br><a href="http://fang/bbs/">返回</a>
</body>
</html>



