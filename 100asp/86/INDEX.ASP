<html>
<head>
<title>网上购物</title>
</head>
<body bgcolor=lightgreen>
<script language="javascript">
<!--
function button1_onclick (){
window.location.href = "register.asp"
}
function button2_onclick (){
window.location.href = "login.asp"
}
-->
</script>
<div align=center>
<table width="500" border="1" name="t1">
<form id=form1 name=form1 action="message.asp"><tr><td colspan=4 align=center><FONT face=华文行楷 size=6>  实例90&nbsp;&nbsp; 网上购物</FONT></td></tr> 
<% 
set rs=server.createobject("adodb.recordset") 
conn = "DBQ=" + server.mappath("mydb.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 
sql="select * from mytable" 
rs.open sql,conn,1,1 
if not rs.EOF or rs.BOF then 
%> 
<tr bgcolor=LawnGreen><td WIDTH=45%>手机型号</td><td WIDTH=25%>生产厂家</td><td WIDTH=20%>价格</td><td WIDTH=15%>数量</td></tr> 
<%
do while not rs.EOF or rs.BOF
%> 
	<tr bgcolor=LightGoldenrodYellow><td><%=rs("name")%></td> 
	<td><%=rs("publish")%></td> 
	<td><%=rs("price")%>元</td> 
	<td><INPUT type="text" name=<%=rs("id")%>></td>
<% 	rs.movenext
	loop 
end if 
set rs=nothing 
set conn=nothing 
%> 
<tr>  
<td colspan=4 align=center>
<INPUT type="button" value="用户注册" name=button1 onclick=button1_onclick()>
<INPUT type="button" value="用户登录" name=button2 onclick=button2_onclick()>
<INPUT type="submit" value="用户提交" id=submit1 name=submit1 onclick=button3_onclick()>
</td> 
</tr></FORM></TABLE> 
</div>