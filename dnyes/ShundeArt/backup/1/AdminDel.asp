<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster")  THEN
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	if request.cookies(Forcast_SN)("ManagePurview")="99999" and request.cookies(Forcast_SN)("purview")="99999" then
		dim rs,tsql
		dim rst
		dim id
		id=request("id")
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql="select * from "& db_Manage_Table &" where "& db_ManageUser_ID &"="&id
		rs.open sql,Conn,3,3
		username=rs(db_ManageUser_Name)
		rs.close
		set rs=nothing
		set rst=server.CreateObject("ADODB.RecordSet")
		if request("name")="del" then
			rst.open "delete from "& db_Manage_Table &" where "& db_ManageUser_ID &"="+request("id"),Conn,3,3
		end if
		rst.close
		set rst=nothing
		conn.close
		set conn=nothing
	else
		Show_Err("对不起，您的管理权限不够！")
		response.end
	end if
	response.redirect "AdminManage.asp"
end if%>