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
		ID=request("ID")
		if ID="" then
			Show_Err("请输入您要删除的ID！<br><br><a href='javascript:history.back()'>回去重来</a>")
			response.end
		end if
		set rs=server.CreateObject ("ADODB.RecordSet")
		rs.Source="select * from "& db_link_Table &" where ID=" & ID
		rs.Open rs.source,conn,1,1
		if rs.EOF then
			Show_Err("非法ID，请确认ID是否存在！<br><br><a href='javascript:history.back()'>回去重来</a>")
			Response.End
		end if
		%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 删除友情链接</title>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<table border="1" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="<%=border%>" align=center>
<form action=LinkDel2.asp method=post>
<tr class="TDtop">
	<td colspan=1 width="100%" height="25" align="center">┊ <strong>删 除 网 站 链 接 确 认</strong> ┊</td>
</tr>
<input type="hidden" name="ID" value="<%=ID%>">
<tr>
	<td width="100%" align=center bgcolor="#FFFFFF">
		<p>&nbsp;</p>
		<p><font color=red>删除网站链接：[ <font color="#000066"> <%=ID%> </font>号 ]？<br><BR>
		（删除后将无法恢复！）</font></p>
		<p>&nbsp;</p>
	</td>
</tr>
<tr class="TDtop">
	<td width="100%"align=center height="25" >
		<input type=submit value="  是  " name="alert_button" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp;&nbsp;&nbsp;&nbsp;
		<input type=submit value="  否  " name="alert_button" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
<tr>
	<td width="100%" align=center bgcolor="#FFFFFF"></td>
</tr>
</form>
</table>
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