<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<%
set rs=server.createobject("ADODB.recordset")
sql="select * from subs where id=" & request("id")
rs.open sql,conn,1,1

Response.ContentType = "img/*"

if request("tu")="tu" then Response.BinaryWrite rs("tu").getChunk(500000)
if request("tu")="stu" then Response.BinaryWrite rs("stu").getChunk(50000)

rs.close
set rs=nothing
set conn=nothing
%>