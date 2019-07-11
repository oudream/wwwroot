<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>实例84：留言本</TITLE>
</HEAD>
<BODY><p align=center><font size=+5>留言本</font></p>
<%
set rs=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("message.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
sql="select * from mytable"
rs.open sql,conn,1,1
if rs.EOF or rs.BOF then
Response.Write "当前没有留言！"
Else
%>
<table>
<%
do while not rs.EOF or rs.BOF
%>
<tr>
<td>
<br>网友&nbsp;&nbsp;<%=rs("user")%>于&nbsp;&nbsp;<%=rs("addtime")%>给&nbsp;&nbsp;<%=rs("towho")%>留言
</td>
</TR>
<tr>
<td>&nbsp;<%=rs("content")%></td>
</tr><%
rs.movenext
loop
%></table><%
end if
set rs=nothing
set conn=nothing
%>
</BODY>
</HTML>
