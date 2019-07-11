<!--#INCLUDE FILE="Conn.asp" -->
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
	dim reglevel
	dim ID
	ID=request("ID")
	reglevel=trim(request("reglevel"))
	
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_User_Table &" where "& db_User_ID &"="&ID
	rs.open sql,ConnUser,3,3
	rs("reglevel")=reglevel
	rs.update
	
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=UserManage.asp"">"
	Show_Message("<p align=center><font color=red>恭喜您!该会员等级设置成功!<br><br>"&freetime&"秒钟后返回上页!</font>")
	response.end
end if%>