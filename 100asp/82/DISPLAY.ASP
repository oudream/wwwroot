<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>计数器</TITLE>
</HEAD>
<BODY>
<%
sql="select * from mytable"
set rs=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("count.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
rs.open sql,conn,1,1
if rs.EOF then
	Response.Write "当前没有留言！"
else%>
<%
do while not rs.EOF%>
<%=rs("user")%>公司的点击率为<%=rs("count")%>
<BR><%
rs.movenext
loop
end if
set rs=nothing
set conn=nothing
%>
</BODY>
</HTML>
