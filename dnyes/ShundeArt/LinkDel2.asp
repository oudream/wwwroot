<%@ Language=VBScript %>
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
		ID=Request.Form("ID")
		button_value=trim(Request.Form("alert_button"))
		if button_value="是" then
			conn.execute("delete from "& db_link_Table &" where ID=" & ID)
		else
			Response.Redirect "linkmanage.asp"
		end if
		response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=LinkManage.asp"">"
		Show_Message("<p align=center><font color=red>ID为["& ID &"]的网站链接友情链接已经删除!<br><br>"&freetime&"秒钟后返回上页!</font>")
		response.end
		rs.close 
		set rs=nothing
		conn.close
		set conn=nothing
	else
			Show_Err("对不起，该功能已经被超级系统管理员关闭，您没有权限操作！")
		response.end
	end if
end if%>