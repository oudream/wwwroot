<html>
<head>
<title>网上书店</title>
</head>
<body bgcolor="#DFEFFF" text="#0000FF">
<table width=500>
<tr><td align=center><font size=+3>网上书店</font></td></tr>
<tr><td><font face="楷体_GB2312" size="+1">请选择您需要的书籍！</font></td></tr>
<tr><td align=right>
<a href="display.asp">查看购物蓝</a>
<a href="clear.asp">清空购物蓝</a></td></tr>
<tr><td>
<table width="90%" border="1" bordercolor="#663399" align="center">
<tr align=center><td width>选购</td>
<td>书名：</td>
<td>出版社</td>
<td>价格</td></tr>
<Form Action=Add.asp Method=POST>
<%
set rs=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("mydb.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
sql="select * from mytable"
rs.open sql,conn,1,1
do while not rs.EOF or rs.BOF
%>
<tr><TD width=10><Input Type=radio Name="CheckID" Value="<%=rs("id")%>"></TD>
<td><%=rs("name")%></td>
<td><%=rs("publish")%></td>
<td><%=rs("price")%>元</td>
<%
rs.movenext
loop
set rs=nothing
set conn=nothing
%>
<tr><td colspan=4 align="center"><INPUT type="submit" value="放入购物蓝"></form>
</td></tr>
</td></tr>
</table>
</table>
</body> 
</html>