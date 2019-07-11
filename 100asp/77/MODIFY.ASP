<script language="javascript">
function confirmDel(id){
window.location.href = "modifydisplay.asp?id=" + id
}
</script>
<%
dim sql
dim rs
set rs=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("mydb.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
sql="select * from mytable"
rs.Open sql,conn,1,1
if rs.EOF then
	Response.Write "NO"
else%>
<P align=center><FONT face=隶书 size=6>实例&nbsp; 同学录</FONT></P>
<table border="1" align=center bordercolor=#3333ff cellspacing="0" cellpadding="0">
<tr><td colspan=7>
<A HREF="index.asp">首页</A>
<A HREF="add.asp">添加</A>
<A HREF="del.asp">删除</A>
修改</td></tr>
<tr align=center>
<td width=80>姓名</td>
<td width=40>性别</td>
<td width=100>生日</td>
<td width=120>电话</td>
<td width=140>家庭住址</td>
<td width=140>所在公司</td>
<td width=50>操作</td>
</tr>
<%
do while not rs.EOF
%>
<TR><TD COLS=5 HEIGHT=5></TD></TR>
<tr>
	<td>&nbsp;<%=rs("username")%></td>
	<td>&nbsp;<%=rs("sex")%></td>
	<td>&nbsp;<%=rs("phone")%></td>
	<td>&nbsp;<%=rs("borndate")%></td>
	<td>&nbsp;<%=rs("address")%></td>
	<td>&nbsp;<%=rs("company")%></td>
	<td>&nbsp;<a href='javascript:confirmDel(<%= rs("id") %>)'>修改</a></td>
</tr>
<%
rs.MoveNext
loop
end if
rs.Close
set rs=nothing
%>
