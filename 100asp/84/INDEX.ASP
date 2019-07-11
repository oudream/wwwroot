<%
const Maxday=22
dim today
dim rs,rstmep
dim sql
dim conn
set rstemp=server.createobject("adodb.recordset")
set rs=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("message.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
%>
<html>
<head>
<title>工资管理</title>
</head>
<body bgcolor="#f0ffff">
<script language="VBScript">
<!--
sub button1_onclick
	location.href="manage.asp?"
	FORM1.submit
end sub
-->
</script>
<P align=center><FONT face=隶书 size=6><STRONG>工资管理</STRONG></FONT> </P>
<table border=0 width=740 align=center>
<tr><td width=160  valign=top>
<table width=100%>
<tr><tr><td>
<FONT size=4 face=隶书>
本月正常上班为<br>
&nbsp;&nbsp;<%response.write maxday%>天<br>
当前日期为：<BR>
<%response.write date()%><p>
<STRONG>说明：</STRONG><BR>
1>周日加班一天计两天工资<br>
2>旷工一天扣除三天工资
</td></tr><tr><td>
<FORM action="manage.asp" id=FORM1 method=post name=FORM1><BR>
请选择人员姓名：<BR>
&nbsp;&nbsp;<SELECT size=1 id=select1 name=select1>
<%
sql = "SELECT * FROM mytable"
rstemp.open sql, conn, 1, 1
do while not rstemp.eof
%><option value=<%=rstemp("id")%>><%= rstemp("name")%></OPTION>
<%
rstemp.movenext
loop
rstemp.close
set rstemp=nothing
%>
</SELECT><BR>
请选择登记情况：<BR>
&nbsp;&nbsp;<SELECT size=1 id=select2 name=select2>
<OPTION value=workday>员工报到</OPTION>
<OPTION value=addday>周日加班</OPTION>
<OPTION value=failday>病假</OPTION>
<OPTION value=thingday>事假</OPTION>
<OPTION value=noworkday>今日旷工</OPTION>
</SELECT><p align=center><INPUT type="button" value="提交" id=button1 name=button1>
</FORM>
</font>
</td></tr></table>
</td>
<td>

<table border="1" width="580">
<%
sql="select * from mytable"
rs.Open sql,conn,1,1
if rs.EOF and  rs.BOF then
	Response.Write "当前没有登记......"
else
%>
<tr><td colspan=8><STRONG>当前共有<%Response.Write rs.RecordCount%>名员工</STRONG></td></tr>
<tr bgcolor=lawngreen>
<TD>员工姓名</TD>
<td>上班天数</td>
<td>周日加班</td>
<td>病假天数</td>
<td>事假天数</td>
<td>旷工天数</td>
<td>基本工资</td>
<td>应得工资</td></tr>
<%
	do while not rs.eof
%>
	<tr bgcolor=wheat>
	<td><%=rs("name")%></td>
	<td><%=rs("workday")%>天</td>
	<td><%=rs("addday")%>天</td>
	<td><%=rs("failday")%>天</td>
	<td><%=rs("thingday")%>天</td>
	<td><%=rs("noworkday")%>天</td>
	<td><%=rs("paymoney")%>元</td> 
	<td><%response.write round(rs("paymoney")*(rs("workday")+rs("addday")*2-rs("noworkday")*3)/maxday,"0")%>元</td>
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