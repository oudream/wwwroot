<html>
<body bgcolor=#c1f7d8>
<center>
<%
dim stroldpassword,strnewpassword,strconfirmpassword
dim strwhere,strsql,strdsn,strchangesql
stroldpassword=request.form("oldpassword")
strnewpassword=request.form("newpassword")
strconfirmpassword=request.form("confirmpassword")
if stroldpassword="" or strnewpassword="" then
    response.write "请输入密码"
    response.end
end if
if strnewpassword<>strconfirmpassword then
    response.write "两次密码不相同"
    response.end
end if
strwhere="where username='" & session("name") & "' and userpassword='" & stroldpassword & "'"
strsql="select * from user " & strwhere
strdsn="dsn=bbs;uid=feng;pwd=feng"
strchangesql="update user set userpassword ='" & strnewpassword & "' " & strwhere
set rs=server.createobject("adodb.recordset")
rs.open strsql,strdsn,1,3
%>
<br>
<%
if rs.recordcount=1 then
   set changers=server.createobject("adodb.recordset")
   changers.open strchangesql,strdsn,1,3
   set changers=nothing
   response.write "密码已成功修改"
else
   response.write "密码输入错误，无法修改密码"
end if
rs.close
set rs=nothing
%>
</body>
</html>