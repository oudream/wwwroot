<% @language=vbscript%>
<html>
<head>
</head>
<body bgcolor=#c1f7d8>
<center>
<%
dim strname,strcontent,strtable,strdsn
strname=request.form("name")
strcontent=request.form("content")
strtoname=request.form("toname")
if trim(strname)="" or trim(strcontent)="" then
    response.write "<br><br>���������ݲ���Ϊ��"
%>
<br><br><a href=javascript:history.bakc()>��һҳ</a><br>
<%
   response.end 
end if
strtable="message"
strdsn="dsn=bbs;uid=feng;pwd=feng"
set rs=server.createobject("adodb.recordset")
rs.open strtable,strdsn,1,3
rs.addnew
rs("messagename")=strname
rs("messagecontent")=strcontent
rs("messagetoname")=strtoname
rs.update
rs.close
set rs=nothing
response.write "ף���㣬�����Ϣ�ɹ��ط���" & strtoname
%>
</center>
</body>
</html>



