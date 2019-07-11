<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE></TITLE>
</HEAD>
<BODY>
<%
dim IID
dim sql
dim rs
dim conn
IID=Request("imgid")
conn="DBQ=" + server.mappath("count.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 
set rs=server.createobject("adodb.recordset")
sql = "UPDATE mytable SET count=count+1 WHERE id = " + cstr(IID)
rs.open sql,conn,1,1
set rs=nothing
set conn=nothing
Response.Redirect "index.asp"
%>
</BODY>
</HTML>
