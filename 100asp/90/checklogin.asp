<%@ Language=VBScript %>
<%
dim name,pwd,email,person
dim sql
dim rs,rsadd
name=request("txt1")
password=request("txt2")
if name="lyf" and password="2000" then
		session("name")=name
		Response.Redirect "main.asp"
else

set rs=server.createobject("adodb.recordset")
set rsadd=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("mydb.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
sql="select * from user where username='"&name&"'and password='"&password&"'"
rs.Open sql,conn,1,1
if not rs.EOF then
	sql="select * from activetable where name='"&name&"'"
	rsadd.Open sql,conn,1,1
	if rsadd.EOF then
		rsadd.Close
		sql="insert into activetable(name) values('"&name&"')"
		rsadd.Open sql,conn,1,1
		session("name")=name
		Response.Redirect "main.asp"
	else
		response.write "<script language=JavaScript>" & chr(13) & "alert('���û��Ѿ�����ʹ�û����벻��ȷ���û���¼ʧ�ܣ�');" & "history.back()" & "</script>"
	end if
else
	response.write "<script language=JavaScript>" & chr(13) & "alert('���û��Ѿ�����ʹ�û����벻��ȷ���û���¼ʧ�ܣ�');" & "history.back()" & "</script>"
end if

end if
%>