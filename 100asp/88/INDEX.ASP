<html>
<head>
<title>图书管理</title>
</head>
<body bgcolor="#f0ffff">
<P align=center><FONT face=隶书 size=6><STRONG>图书管理</STRONG></FONT> </P>
<table border=0 width=720 align=center>
<tr><td width=120  valign=top><!--#include file="left.asp"-->
</td>
<td>
<table border="1" width="600">
<%
dim rs
dim sql
dim conn
	dim author
	author=request("author")
	set rs=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("book.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
if request("author")<>"" then
	sql="select * from mytable where author like'%" & author & "%'"
else
	sql="select * from mytable"
end if
rs.Open sql,conn,1,1
totalput=rs.RecordCount
if rs.EOF and  rs.BOF then
	Response.Write "当前没有图书......"
else
%><tr><td colspan=5>符合条件的共有<%Response.Write totalput%>本图书</td></tr>
<tr bgcolor=lawngreen><TD>图书编号</TD><td>书名</td><td>作者</td><td>是否在架</td><td>借阅人</td></tr>
<%
	do while not rs.eof
%>
	<tr bgcolor=wheat>
	<td><%=rs("id")%></td>
	<td>《<%=rs("name")%>》</td>
	<td><%=rs("author")%></td>
	<td><%if rs("islend")= true then
		Response.Write "已经借出"
		else
		Response.Write "在架"
		end if%>
	</td>
	<td>&nbsp;<%=rs("username")%></td>
</tr>
<%
	rs.movenext
	loop
end if
rs.Close 
set rs=nothing
set conn=nothing
%>
</table>
</td></tr></table>
</body>
</html>