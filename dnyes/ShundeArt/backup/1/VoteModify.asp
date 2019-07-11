<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp" -->
<%
IF not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster" or request.cookies(Forcast_SN)("ManageKEY")="check" or request.cookies(Forcast_SN)("ManageKEY")="typemaster") THEN
	Show_Err("对不起，您的管理权限不够！")
	response.end
else
	if votemana="1" or (request.cookies(Forcast_SN)("ManagePurview")="99999" and request.cookies(Forcast_SN)("purview")="99999") then
		id=request.QueryString("id")
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_Vote_Table &" where id="&id
		rs.open sql,conn,1,1
		if rs.eof then
			Show_Err("<li>操作错误！请联系管理员<br><br>－－－<a href='javascript:history.back()'>回去重来</a>－－－")
			Response.End 
		else
			Title=rs("Title")
			DateAndTime=rs("DateAndTime")
		end if
		%>
<head><head>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 投票管理</title>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<table border="1" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="<%=border%>" align=center>
<form method="POST" action="VoteSave.asp?id=<%=id%>">
<tr class="TDtop">
	<td width="100%" height="25" colspan=2 align=center>┊ <strong>修 改 网 站 投 票</strong> ┊</td>
</tr>
<tr>
	<td colspan=2 align=center></td>
</tr>
<tr>
	<td width="40%" align="right">主题：</td>
	<td width="60%"><input type="text" name="Title" value="<%=Title%>" size="20" style="font-family: 宋体; font-size: 9pt"></td>
</tr>
		<%for i=1 to 8%>
<tr>
	<td align="right">选项<%=i%>：</td>
	<td>
		<input type="text" name="select<%=i%>" value="<%=rs("select"&i)%>" size="20" style="font-family: 宋体; font-size: 9pt"> 票数：<input type="text" name="answer<%=i%>" value="<%=rs("answer"&i)%>" size="5" style="font-family: 宋体; font-size: 9pt">
	</td>
</tr>
		<%next%>
<tr>
	<td align="right">发布时间：</td>
	<td>
		<input type="text" name="dat" value="<%=DateAndTime%>" disabled size="20" style="font-family: 宋体; font-size: 9pt">
	</td>
</tr>
<tr>
	<td colspan=2 align=center></td>
</tr>
<tr class="TDtop">
	<td colspan=2 align=center>
		<input type="hidden" value="edit" name="act">
		<input type="button" value=" 返 回 " onclick="javascript:history.go(-1)" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp; 
		<input type="submit" value=" 修 改 " name="cmdok" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp; 
		<input type="reset" value=" 重 置 "  name="cmdcancel" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
<tr>
	<td colspan=2 align=center></td>
</tr>
</form>
</table>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
	<%else
		Show_Err("对不起，该功能已经被超级系统管理员关闭，您没有权限操作！")
		response.end
	end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if%>