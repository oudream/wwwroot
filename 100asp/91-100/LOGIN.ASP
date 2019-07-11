<%@language=vbscript%>
<%response.buffer=true%>
<html>
<body bgcolor=#c1f7d8>
<center>
<%
dim strname,strpassword,sql
strname=request.form("Name")
strpassword=request.form("password")
sql="select * from user where username='" & strname & "'"
if strname="" then
  response.write "帐号不能为空<p></p>"
%>
<a href=javascript:history.back()>上一页</a>
<%
  response.end
end if
if strpassword="" then
  response.write "密码不能为空<p></p>"
%>
<a href=javascript:history.back()>上一页</a>
<%
  response.end
end if
set rstemp=server.createobject("adodb.recordset")
rstemp.open sql,"dsn=bbs;uid=feng;pwd=feng"
if rstemp.RecordCount=1 then
  if rstemp("userpassword")= strpassword then
  session("name")=strname
  response.redirect "user.htm"
  else
  response.write "密码错误<br><br>"
  end if
else
  response.write "帐号错误<br><br>"
%>
<a href=javascript:history.bakc()>上一页</a>
<%
  response.end
end if
rstemp.close
set rstemp=nothing
%>
</body>
</html>

