<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>查看购物蓝</TITLE>
</HEAD>
<BODY text=MediumSlateBlue>
<table width=70% align=center>
<tr><td><font size=+2><STRONG><U>查看购物蓝</U></STRONG></font></td></tr>
<%
set rs=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("mydb.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
sql="select * from shop"
rs.open sql,conn,1,1
if rs.EOF or rs.BOF then
	response.write "<script language=JavaScript>" & chr(13) & "alert('你的购物蓝是空的！');" & "history.back()" & "</script>" 
Else
%>
<tr bgcolor=LawnGreen><td>你购买的图书</td><td>图书的价格</td><td>购买的时间</td></tr>
<%
do while not rs.EOF or rs.BOF
%>
	<tr bgcolor=Aquamarine><td><%=rs("name")%></td>
	<td><%=rs("price")%>元</td>
	<td><%=rs("addtime")%></td>
	<%
	rs.movenext
	loop
end if
set rs=nothing
set conn=nothing
%>
</table>
</body>
</html>