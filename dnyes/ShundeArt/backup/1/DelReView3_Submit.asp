<%@ Language=VBScript %>
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
	ReViewID=Request.Form("ReViewID")
	button_value=trim(Request.Form("alert_button"))
	if button_value="确定" then
		conn.execute("delete from "& db_Review_Table &" where ReViewID=" & ReViewID)
	else
		Response.Redirect "ReViewManage.asp"
	end if

	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=ReViewManage.asp"">"
	Show_Message("评论 ID =["& ReViewID &"]删除成功！!<br><br>"&freetime&"秒钟后返回上页!</font>")
	response.end
end if
%>