<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="char.inc"-->
<!--#include file="config.asp"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp" -->
<!--#include file="ChkManage.asp" -->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" and request.cookies(Forcast_SN)("ManagePurview")="99999") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 添加管理用户</title>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<link rel="stylesheet" type="text/css" href="site.css">
<script LANGUAGE="javascript">
<!--
function FrmAddLink_onsubmit() {
if (document.FrmAddLink.username.value=="")
	{
	  alert("对不起，请输入管理用户名！")
	  document.FrmAddLink.username.focus()
	  return false
	 }
else if (document.FrmAddLink.username.value.length < 2)
	{
	  alert("管理用户名能不能长一点！")
	  document.FrmAddLink.username.focus()
	  return false
	 }
else if (document.FrmAddLink.username.value.length > 30)
	{
	  alert("管理用户名太长了吧！")
	  document.FrmAddLink.username.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd.value=="")
	{
	  alert("对不起，请输入密码！")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd.value.length < 4)
	{
	  alert("为了安全，密码应该长一点！")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd.value.length > 16)
	{
	  alert("密码太长了吧！")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.username.value==document.FrmAddLink.passwd.value)
	{
	  alert("为了安全，管理用户名与密码不应该相同！")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd2.value=="")
	{
	  alert("对不起，请您输入验证密码！")
	  document.FrmAddLink.passwd2.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd2.value !== document.FrmAddLink.passwd.value)
	{
	  alert("对不起，您两次输入的密码不一致！")
	  document.FrmAddLink.passwd2.focus()
	  return false
	 }
else if (document.FrmAddLink.fullname.value=="")
	{
	  alert("对不起，请输入该管理用户的真实姓名！")
	  document.FrmAddLink.fullname.focus()
	  return false
	 }
else if (document.FrmAddLink.depid.value=="")
	{
	  alert("对不起，请选择该管理用户的工作单位！")
	  document.FrmAddLink.depid.focus()
	  return false
	 }
else if (document.FrmAddLink.select.value=="")
	{
	  alert("对不起，请选择该管理用户的权限！")
	  document.FrmAddLink.select.focus()
	  return false
	 }
}

//-->
</script>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<div align="center">
<table align=center border="1" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="<%=border%>">
<form align="center" method="post" name="FrmAddLink" LANGUAGE="javascript" onsubmit="return FrmAddLink_onsubmit()" action="AdminAdd2.asp">
<tr>
	<td width="100%" height=32 colspan="2" class="TDtop"><p align="center" >┊<strong> 添 加 新 管 理 用 户 </strong>┊</td>
</tr>
<tr> 
	<td width="46%" height="32" align="right" valign="middle"> 
		<div align="center"><span class="smallFont">管 理 用 户 名：</span></div>
	</td>
	<td width="54%" height="32"> 
		<div align="left"> <input name="username" size="26" class="smallInput" maxlength="30" style="font-family: 宋体; font-size: 9pt" ></div>
	</td>
</tr>
<tr>
	<td width="46%" height="32" align="right"> 
		<div align="center"><span class="smallFont">密&nbsp;&nbsp;&nbsp; 码：</span></div>
	</td>
	<td width="54%" height="32"> 
		<div align="left"><input name="passwd" size="26" class="smallInput" maxlength="15" style="font-family: 宋体; font-size: 9pt" type="password" ></div>
	</td>
</tr>
<tr>
	<td width="46%" height="32" align="right"> 
		<div align="center"><span class="smallFont">验证密码：</span></div>
	</td>
	<td width="54%" height="32"> 
		<div align="left"><input name="passwd2" size="26" class="smallInput" maxlength="15" style="font-family: 宋体; font-size: 9pt" type="password" ></div>
	</td>
</tr>
<tr> 
	<td width="46%" height="32" align="right"> 
		<div align="center"><span class="smallFont">真实姓名：</span></div>
	</td>
	<td width="54%" height="32"> 
		<div align="left"><input name="fullname" size="26" class="smallInput" maxlength="10" style="font-family: 宋体; font-size: 9pt" ></div>
	</td>
</tr>
<tr> 
	<td width="46%" height="32" align="right">
		<div align="center"><span class="smallFont">对应前台普通发文用户：</span></div>
	</td>
	<td width="54%" height="32"> 
		<div align="left"><input name="adder" size="26" class="smallInput" maxlength="10" style="font-family: 宋体; font-size: 9pt" >（填错将使本超管用户无法登陆）</div>
	</td>
</tr>
<tr>
	<td width="46%" height="32" align="right"> 
		<div align="center"><span class="smallFont">工作单位：</span></div>
	</td>
	<td width="54%" height="32"> <% 
	set rstype=createobject("adodb.recordset")
	sql="select * from "& db_Dep_Table &" order by id"
	rstype.Open sql,conn,1,3
	%>
		<select name="depid" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
			<option value="">请选择工作单位</option>
	<%do while not rstype.EOF%>
			<option value="<%=rstype("id")%>"><%=rstype("depname")%>==<%=rstype("deptype")%></option>
		<%
		rstype.MoveNext
	loop
	set rstype=nothing
	%>
		</select>
	</td>
</tr>
<tr> 
	<td width="46%" height="32" align="right"> 
		<div align="center">管理用户权限：</div>
	</td>
	<td width="54%" height="32">
		<select name="purview" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
			<option  selected value="">请选择用户权限</option>
			<option value=99999 selected>超级管理员</option>
			<option value=1>系统管理员</option>
		</select>
	</td>
</tr>
<tr> 
	<td width="46%" colspan="2" height="32" align="right">&nbsp;</td>
</tr>
<tr>
	<td width="100%" colspan="2" height=32><p align="center">
		<input name="shenhe" type="hidden" value="0">
		<input name="oskey" type="hidden" value="super">
		<input type="button" value=" 返 回 " onclick="javascript:history.go(-1)" class="unnamed5"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp;           
		<input type="submit" value=" 确 定 " name="cmdOk" class="buttonface" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp; 
		<input type="reset" value=" 重 填 " name="cmdReset" class="buttonface" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
</form>
</table>
</div>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
<%END IF%>
