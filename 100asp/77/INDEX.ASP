<%@ Language=VBScript %>
<HTML>
<HEAD>
<title>同学录</title>
</HEAD>
<BODY background=background.gif>
<P align=center><FONT face=隶书 size=6>实例&nbsp; 同学录</FONT></P>

<table border="1" align=center bordercolor=#0099cc cellspacing="0" cellpadding="0">
<tr><td colspan=6><A HREF="add.asp">添加</A>
<A HREF="del.asp">删除</A>
<A HREF="modify.asp">修改</A></td></tr>
<% 
set rs=server.createobject("adodb.recordset") 
conn = "DBQ=" + server.mappath("mydb.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 
sql="select * from mytable" 
rs.open sql,conn,1,1 
if rs.EOF or rs.BOF then
	Response.Write "数据库中暂无资料!"
else
do while not rs.EOF or rs.BOF
%>
<tr>
	<td colspan=2 align=center><FONT face=隶书 size=4><STRONG><%=rs("username")%></STRONG></font></td>
	<td>性别：</td><td>&nbsp;<%=rs("sex")%></td>
	<td>生日：</td><td>&nbsp;<%=rs("borndate")%></td>
</tr>
<tr>
	<td width=100>联系电话：</td><td width=120>&nbsp;<%=rs("phone")%></td>
	<td width=100>呼机号：</td><td width=120>&nbsp;<%=rs("BP")%></td>
	<td width=100>手机号：</td><td width=120>&nbsp;<%=rs("policy")%></td>
</tr>
<tr>
	<td>所在公司：</td><td>&nbsp;<%=rs("company")%></td>
	<td>电子信箱：</td><td>&nbsp;<%=rs("email")%></td>	
	<td>家庭住址：</td><td>&nbsp;<%=rs("address")%></td>
</tr>
<tr><td colspan=6 height=10></td></tr>
<%
rs.movenext
loop
end if
set rs=nothing 
set conn=nothing
%>
</table>
</BODY>
</HTML>
