<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp"-->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	if request.cookies(Forcast_SN)("Purview")="99999" and request.cookies(Forcast_SN)("KEY")="super" then
		dim rs,tsql
		dim rst
		dim id
		id=request("id")
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql="select * from "& db_User_Table &" where "& db_User_ID &"="&id
		rs.open sql,ConnUser,3,3
		username=rs(db_User_Name)
		rs.close
		set rs=nothing
		set rst=server.CreateObject("ADODB.RecordSet")
		if request("name")="del" then
			rst.open "delete from "& db_User_Table &" where "& db_User_ID &"="&request("id"),ConnUser,3,3
		end if
		response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=UserManage.asp"">"
		Show_Message("<p align=center><font color=red>用户ID ["&request("id")&"] 删除成功!<br><br>"&freetime&"秒钟后返回上页!</font>")
		response.end
	else
		Show_Err("对不起，您无权进行用户的删除！")
		response.end
	end if
end if%>