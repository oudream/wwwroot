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
  response.write "�ʺŲ���Ϊ��<p></p>"
%>
<a href=javascript:history.back()>��һҳ</a>
<%
  response.end
end if
if strpassword="" then
  response.write "���벻��Ϊ��<p></p>"
%>
<a href=javascript:history.back()>��һҳ</a>
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
  response.write "�������<br><br>"
  end if
else
  response.write "�ʺŴ���<br><br>"
%>
<a href=javascript:history.bakc()>��һҳ</a>
<%
  response.end
end if
rstemp.close
set rstemp=nothing
%>
</body>
</html>

