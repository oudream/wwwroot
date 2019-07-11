<%@ Language=VBScript %>
<%
dim name,pwd,email,person
dim sql
dim rs
name=request("txt1")
password=request("txt2")
set rs=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("mydb.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
sql="select * from user where username='"&name&"'and password='"&password&"'"
Response.Write sql
rs.Open sql,conn,1,1
if not rs.EOF then
	session("name")=name
	Response.Redirect "index.asp"
else
	Response.Write "ÓÃ»§µÇÂ¼Ê§°Ü!!"
end if
%>