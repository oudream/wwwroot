<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	ReViewID=request("ReViewID")
	if ReViewID="" then
		Show_Err("请输入您要删除的评论ID!<br><br>－－－<a href='javascript:history.back()'>回去重来</a>－－－")
		response.end
	end if
	set rs=server.CreateObject ("ADODB.RecordSet")
	rs.Source="select * from "& db_Review_Table &" where ReViewID=" & ReViewID
	rs.Open rs.source,conn,1,1
	if rs.EOF then
		Show_Err("非法评论ID，请确认ReviewID是否存在!<br><br>－－－<a href='javascript:history.back()'>回去重来</a>－－－")
		Response.End
	end if
	%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 删除评论</title>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<div align="center">
<table border="1" width="100%" cellspacing="0" bgcolor="#000000" cellpadding="0" style="border-collapse: collapse" bordercolor="<%=border%>">
<form action=DelReview3_Submit.asp method=post>
<tr class="TDtop">
	<td height=25 class="TDtop" align="center">┊ <B>删 除 确 认</B> ┊ </td>
</tr>
<input type="hidden" name="ReViewID" value="<%=ReViewID%>">
<tr>
	<td width="100%" align=center bgcolor="#FFFFFF">
		<p>　</p>
		<p>
			<font color=red>
				删除评论（留言）：[ <font color="#000066"> <%=ReViewID%> </font>号 ]？<br><BR>
				（删除后将无法恢复！）
			</font>
		</p>
		<p></p>
	</td>
</tr>
	<%if (delreview="1" and request.cookies(Forcast_SN)("ManageKEY")="input") or (shdelreview="1" and request.cookies(Forcast_SN)("ManageKEY")="check") or request.cookies(Forcast_SN)("ManageKEY")="super" then%>
<tr class="TDtop">
	<td height=25 class="TDtop" align="center">
		<input type=submit value="确定" name="alert_button" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp;&nbsp;&nbsp;&nbsp;
		<input type=button value="取消" name="alert_button" class=button onClick="javascript:history.go(-1)" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
	<%end if%>
</form>
</table>
</div>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
<%end if%>