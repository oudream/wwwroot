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
		sql="delete from "& db_Vote_Table &" where ID="&id
		rs.open sql,conn,1,3
		conn.close
		set conn=nothing
		response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=VoteManage.asp"">"
		Show_Message("<p align=center><font color=red>ID号为["& id &"]的投票主题删除成功!<br><br>"&freetime&"秒钟后返回上页!</font>")
		response.end
	else
		Show_Err("对不起，该功能已经被超级系统管理员关闭，您没有权限操作！")
		response.end
	end if
end if
%>