<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster" or request.cookies(Forcast_SN)("ManageKEY")="check" or request.cookies(Forcast_SN)("ManageKEY")="typemaster") then
	Show_Err("对不起，您的管理权限不够！")
	response.end
else
	if linkmana="1" or (request.cookies(Forcast_SN)("ManagePurview")="99999" and request.cookies(Forcast_SN)("purview")="99999") then
		%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 友情链接修改</title>
<script LANGUAGE="javascript">
<!--
function FrmAddLink_onsubmit() {
 if (document.FrmAddLink.linktype.value=="")
	{
	  alert("对不起，请选择您申请的友情链接的类型！")
	  document.FrmAddLink.linktype.focus()
	  return false
	 }
else if (document.FrmAddLink.webname.value=="")
	{
	  alert("对不起，请输入您的网站名称！")
	  document.FrmAddLink.webname.focus()
	  return false
	 }
else if (document.FrmAddLink.weburl.value=="")
	{
	  alert("对不起，请输入您的网站地址！")
	  document.FrmAddLink.weburl.focus()
	  return false
	 }
else if (document.FrmAddLink.weburl.value=="http://")
	{
	  alert("对不起，请输入您的网站地址！")
	  document.FrmAddLink.weburl.focus()
	  return false
	 }
else if (document.FrmAddLink.webmaster.value=="")
	{
	  alert("对不起，不知您是哪位！")
	  document.FrmAddLink.webmaster.focus()
	  return false
	 }
else if (document.FrmAddLink.content.value=="")
	{
	  alert("对不起，请输入您的网站的简单介绍！")
	  document.FrmAddLink.content.focus()
	  return false
	 }
else if (document.FrmAddLink.email.value=="")
	{
	  alert("对不起，请输入您的电子邮件！")
	  document.FrmAddLink.email.focus()
	  return false
	 }
else if (document.FrmAddLink.email.value=="@")
	{
	  alert("对不起，请输入您的电子邮件！")
	  document.FrmAddLink.email.focus()
	  return false
	 }
}

//-->
</script>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
		<%
		dim sql
		dim rs
		sql="select * from "& db_link_Table &" where id="&request("id")
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		%> 
<div align="center">
<table border="1" width="100%" cellpadding="0" cellspacing="0" align="center" bordercolor=<%=border%> style="border-collapse: collapse">
<form align="center" method="post" name="FrmAddLink" LANGUAGE="javascript" onsubmit="return FrmAddLink_onsubmit()" action="LinkSaveEdit.asp?id=<%=request("id")%>">
<tr class="TDtop">
	<td colspan=2 width="100%" height="25" align="center">┊ <strong>友 情 链 接 修 改</strong> ┊</td>
</tr>
<tr>
	<td colspan=2 width="100%" height="25"></td>
</tr>
<tr> 
<tr>
	<td width="45%" height="32" align="right"> 
		<span class="smallFont">类&nbsp;&nbsp;&nbsp; 型：</span>
	</td>
	<td width="55%" height="32"> 
		<select size="1" name="linktype" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
			<option <%if rs("linktype")="2" then%>selected<%end if%> value="2">首页LOGO链接</option>
			<option <%if rs("linktype")="1" then%>selected<%end if%> value="1">其他LOGO链接</option>
			<option  <%if rs("linktype")="0" then%>selected<%end if%> value="0">文字链接</option>
		</select>
	</td>
</tr>
<tr> 
	<td height="32" align="right" valign="middle"> 
		<span class="smallFont">网站名称：</span>
	</td>
	<td height="32"> 
		<div align="left">
			<input name="webname" size="26" class="smallInput" maxlength="20" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="这里请输入您的网站名称，最多为20个汉字" value="<%=rs("webname")%>">
		</div>
	</td>
</tr>
<tr>
	<td height="32" align="right"> 
		<span class="smallFont">网站地址：</span>
	</td>
	<td height="32"> 
		<div align="left"> 
			<input name="weburl" size="26" class="smallInput" maxlength="100" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"type="text" value="<%=rs("weburl")%>" title="这里请输入您的网站地址，最多为50个字符，前面必须带http://">
		</div>
	</td>
</tr>
<tr>
	<td height="32" align="right"> 
		<span class="smallFont">网站LOGO：</span>
	</td>
	<td height="32"> 
		<div align="left"> 
			<input name="logo" size="26" class="smallInput" maxlength="100" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"type="text"  value="<%=rs("logo")%>" title="这里请输入您的网站LOGO地址，最多为50个字符，如果您在第一选项选择的是文字链接，这项就不必填">
		</div>
	</td>
</tr>
<tr>
	<td height="32" align="right"> 
		<span class="smallFont">站&nbsp;&nbsp;&nbsp;长：</span>
	</td>
	<td height="32"> 
		<div align="left"> 
			<input name="webmaster" size="26" class="smallInput" maxlength="20" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"type="text"  title="这里请输入您的大名了，不然我知道您是谁啊。最多为20个字符" value="<%=rs("webmaster")%>">
		</div>
	</td>
</tr>
<tr> 
	<td height="32" align="right"> 
		<span class="smallFont">电子邮件：</span>
	</td>
	<td height="32"> 
		<div align="left"> 
			<input name="email" size="26" class="smallInput" maxlength="30" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"type="text"  value="<%=rs("email")%>" title="这里请输入您的联系电子邮件，最多为30个字符">
		</div>
	</td>
</tr>
<tr>
	<td height="32" align="right"> 
		<span class="smallFont">网站简介：</span>
	</td>
	<td height="32" valign="middle"> 
		<textarea rows="5" name="content" cols="29" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="这里请输入您的网站的简单介绍"><%=rs("content")%></textarea>
	</td>
</tr>
<tr>
	<td colspan=2 width="100%" height="25"></td>
</tr>
<tr class="TDtop">
	<td colspan=2 height="32" align="center"> 
		<div align="center">
			<input type="hidden" name="pass" value="1">
			<input type="submit" value=" 确 定 " name="cmdOk" class="buttonface" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp;
			<input type="reset" value=" 重 置 " name="cmdReset" class="buttonface" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		</div>
	</td>
</tr>
</form>
</table>
</center>
</div>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
		<%rs.close 
		set rs=nothing
		conn.close
		set conn=nothing
	else
			Show_Err("对不起，该功能已经被超级系统管理员关闭，您没有权限操作！")
		response.end
	end if
end if%>