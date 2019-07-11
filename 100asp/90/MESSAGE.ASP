<P align=center><A href="message.asp">刷新人员列表</A><P>
<%@ Language=VBScript %>
<%
set rs=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("mydb.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
sql="select * from activetable"
rs.Open sql,conn,1,1
Response.Write "当前在线人员：<br>"
do while not rs.eof 
Response.Write  rs("name")&"<br>"
rs.MoveNext
loop
set rs=nothing
%>
</P>
