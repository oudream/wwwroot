<% @language=vbscript%>
<html>
<head>
</head>
<body bgcolor=#c1f7d8>
<center>
<%
dim strName,strEmail,strHomepage,strContent,strTable,strDsn
strName=Request.Form("name")
strEmail=Request.Form("email")
strHomepage=request.form("homepage")
strcontent=request.form("content")
strtoname=request.form("toname")
if trim(strname)="" or trim(strcontent)="" then
   response.write "<br><br>���������ݲ���Ϊ��"
%>
<br><b><a href=javascript:hisory.back()>��һҳ</a><br>
<%
   response.end
end if
strtable="message"
strDsn="dsn=bbs;uid=feng;pwd=feng"
set rs=server.createobject("adodb.recordset")
rs.open strtable,strdsn,1,3
rs.addnew
rs("messagename")=srname
rs("messageemail")=stremail
rs("messagehomepage")=strhomepage
rs("messagecontent")=strcontent
rs("messagetoname")=strtoname
rs.update
rs.close
set rs=nothing
response.write "ף���㣬�����Ϣ�ɹ��ط����˹���Ա"
%>
</center>
</body>
</html>
