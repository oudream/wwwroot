<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 添加用户</title>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<link rel="stylesheet" type="text/css" href="site.css">
<script LANGUAGE="javascript">
<!--
function FrmAddLink_onsubmit() {
if (document.FrmAddLink.username.value=="")
	{
	  alert("对不起，请输入用户名！")
	  document.FrmAddLink.username.focus()
	  return false
	 }
else if (document.FrmAddLink.username.value.length < 2)
	{
	  alert("用户名能不能长一点！")
	  document.FrmAddLink.username.focus()
	  return false
	 }
else if (document.FrmAddLink.username.value.length > 30)
	{
	  alert("用户名太长了吧！")
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
	  alert("为了安全，用户名与密码不应该相同！")
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
	  alert("对不起，请输入该用户的真实姓名！")
	  document.FrmAddLink.fullname.focus()
	  return false
	 }
else if (document.FrmAddLink.depid.value=="")
	{
	  alert("对不起，请选择该用户的工作单位！")
	  document.FrmAddLink.depid.focus()
	  return false
	 }
else if (document.FrmAddLink.select.value=="")
	{
	  alert("对不起，请选择该用户的权限！")
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
<table align=center border="1" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" >
<form align="center" method="post" name="FrmAddLink" LANGUAGE="javascript" onsubmit="return FrmAddLink_onsubmit()" action="UserAdd2.asp">
<tr>
	<td width="100%" height=32 colspan="2" class="TDtop">
		<p align="center" >┊<strong> 添 加 新 用 户 </strong>┊
	</td>
</tr>
<tr> 
	<td width="46%" height="32" align="right" valign="middle"> 
		<div align="center"><span class="smallFont">用 户 名：</span></div>
	</td>
	<td width="54%" height="32"> 
		<div align="left">
			<input name="username" size="26" class="smallInput" maxlength="30" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		</div>
	</td>
</tr>
<tr>
	<td height="32" align="right"> 
		<div align="center"><span class="smallFont">密&nbsp;&nbsp;&nbsp;码：</span></div>
	</td>
	<td height="32"> 
		<div align="left"> 
			<input name="passwd" size="26" class="smallInput" maxlength="15" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" type="password" >
		</div>
	</td>
</tr>
<tr>
	<td height="32" align="right"> 
		<div align="center"><span class="smallFont">验证密码：</span></div>
	</td>
	<td height="32"> 
		<div align="left"> 
			<input name="passwd2" size="26" class="smallInput" maxlength="15" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" type="password" >
		</div>
	</td>
</tr>
<tr> 
	<td height="32" align="right"> 
		<div align="center"><span class="smallFont">真实姓名：</span></div>
	</td>
	<td height="32"> 
		<div align="left"> 
			<input name="fullname" size="26" class="smallInput" maxlength="10" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		</div>
	</td>
</tr>
<tr>
	<td height="32" align="right"> 
		<div align="center"><span class="smallFont">工作单位：</span></div>
	</td>
	<td height="32">
	<% 
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
	<td height="32" align="right"><div align="center">用户权限：</div></td>
	<td height="32">
		<select name="oskey" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<option  selected value="">请选择用户权限</option>
	<%if request.cookies(Forcast_SN)("key")="super" then%>
		<option value="super">系统管理员</option>
		<option value="check">文章审核员</option>
		<option value="typemaster">总栏管理员</option>
	<%end if%>
	<%if request.cookies(Forcast_SN)("key")="typemaster" or request.cookies(Forcast_SN)("key")="super" then%>
		<option value="bigmaster">大类管理员</option>
	<%end if%>
		<option value="smallmaster">小类管理员</option>
	<%if request.cookies(Forcast_SN)("key")="super" then%>
		<option value="selfreg">注册用户</option>
	<%end if%>
		</select>
	</td>
</tr>
<tr> 
	<td height="32" align="right"> 
		<div align="center">是否需要审核：<br>该功能仅对小类管理员和注册用户有效！</div>
	</td>
	<td height="32">
		<select name="shenhe" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<option selected value="">该用户需要审核吗</option>
		<option value="1">需要</option>
		<option value="0">不需要</option>
		</select>默认值为需要审核
	</td>
</tr>
<tr>
	<td width="100%" colspan="2" height=32>
		<p align="center">
		<input name="purview" type="hidden" value="1">
		<input name="adder" type="hidden" value="<%=request.cookies(Forcast_SN)("username")%>">
		<input type="button" value=" 返 回 " onclick="javascript:history.go(-1)" class="unnamed5" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp;
		<input type="submit" value=" 确 定 " name="cmdOk" class="buttonface" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp;
		<input type="reset" value=" 重 置 " name="cmdReset" class="buttonface" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
</form>
</table>
</center>
</div>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
<%end if%>