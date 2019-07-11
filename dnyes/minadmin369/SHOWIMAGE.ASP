<!--#include file="SHEI45FECSA.ASP" -->
<%
set rs=server.createobject("ADODB.recordset")
sql="select * from subs where id=" & request("id")
rs.open sql,conn,1,1

Response.ContentType = "img/*"

if request("tu")="tu" then Response.BinaryWrite rs("tu").getChunk(7500000)
if request("tu")="btu" then Response.BinaryWrite rs("btu").getChunk(7500000)
if request("tu")="stu" then Response.BinaryWrite rs("stu").getChunk(7500000)

rs.close
set rs=nothing
conn.close
set conn=nothing
%>
