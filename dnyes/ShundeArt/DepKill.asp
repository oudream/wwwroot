<%@ Language=VBScript %>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp"-->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<%
IF request.cookies(Forcast_SN)("ManageKEY")<>"super" then
	Show_Err("对不起，您的后台管理权限不够！")
	response.end
else
	dim ID
	dim depName
	ID = Request("ID")
	depName=request("depName")
	%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 单位删除</title>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<table border="1" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="<%=border%>" align=center>
<form action=depKillok.asp method=post>
<tr>
	<td width="100%" align="center" height="55" bgcolor="<%=m_top%>"><b>删 除 单 位 确 认</b></td>
</tr>
<input type="hidden" name="id" value="<%=id%>">
<tr>
	<td width="100%" align=center bgcolor="#FFFFFF"> 
		<p>　</p>
		<p>删除单位：[ <%=depName%> ]？<font color=red><br><br>（此操作将无法恢复！）</font></p>
		<p>　</p>
	</td>
</tr>
<tr>
	<td width="100%"align=center bgcolor="<%=m_top%>" height="55" > 
		<input type=submit value="确定" name="alert_button" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp;&nbsp;&nbsp;&nbsp;
		<input type=submit value="取消" name="alert_button" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
</form>
</table>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
<%end if%>