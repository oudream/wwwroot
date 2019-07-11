<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->

<%
IF request.cookies(Forcast_SN)("ManageKEY")<>"super" then
	Show_Err("对不起，您的后台管理权限不够！")
	response.end
else
	Set rs6 = Server.CreateObject("ADODB.Recordset")
	sql6 ="select * from "& db_dep_Table &" where DepName='" & Request("depname") & "' and Deptype='" & Request("deptype") & "' order by ID"
	rs6.open sql6,Conn,3,3
	%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 修改部门</title>
<LINK href=site.css rel=stylesheet>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<form method="POST" action="depEditok.asp">
<tr>
	<td width="100%" height="25" class="TDtop"> 
		<p align="center" >┊ <strong>修 改 单 位 部 门</strong> ┊
	</td>
</tr>
<tr align="center"> 
	<td width="100%" bgcolor="#FFFFFF" height="150">
			<input type="hidden" name="name" value='<%=request("depName")%>'>
			<input type="hidden" name="ID" value='<%=request("ID")%>'>
			<input type="hidden" name="type" value='<%=request("deptype")%>'>
		<table border="0" cellspacing="1" width="100%" height="70">
		<tr> 
			<td align="right" height="25" width="50%">单位名称：</td>
			<td align="left" width="50%">
				<input type="text" name="depname" size="20" maxlength="20" value='<%=request("depName")%>' style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
			</td>
		</tr>
		<tr>
			<td align="right" height="25">部门名称：</td>
			<td align="left" height="25">
				<input type="text" name="deptype" size="20" maxlength="20" value='<%=request("deptype")%>' style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center"> 
	<td width="100%" height="60" class="TDtop1">
		<input type="submit" value=" 修 改 " name="cmdok" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp;
		<input type="reset" value=" 复 原 " name="cmdcancel" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp; 
		<input type="button" name="ok" value=" 放 弃 " onClick="javascript:history.go(-1)" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
</form>
</table>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
	<%rs6.close
	set rs6=nothing
	conn.close
	set conn=nothing
end if%>