<%@ Language=VBScript %>
<%Response.Buffer=true%>
<html>
<head>
<META HTTP-EQUIV="REFRESH" CONTENT=5;url="content.asp">
</head>
<body>
<%
set rs=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("mydb.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
sql="select * from message order by id desc"
rs.Open sql,conn,1,1
do while not rs.eof 
Response.Write  rs("name")&"หตฃบ"&rs("message")&"<br>"
rs.MoveNext
loop
set rs=nothing
%>
</body>