<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="chkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	SpecialID = Request("SpecialID")
	dim Specialname
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from "& db_Special_Table &" where SpecialID="&SpecialID
	rs.open sql,conn,3,3
	Specialname=rs("Specialname")
	rs.close
	set rs=nothing
	
	button_value=trim(Request.Form("alert_button"))
	if button_value="确定" then
		conn.execute("delete from "& db_Special_Table &" where Specialid=" &Specialid)
		if killnews=1 then
			conn.execute("delete from "& db_News_Table &" where Specialid='" &Specialid &"'")
		end if
		conn.close
		set conn=nothing
		response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=SpecialManage.asp"">"
		Show_Message("<p align=center><font color=red>恭喜您!您选择的专题已经被删除!<br><br>"&freetime&"秒钟后返回上页!</font>")
		response.end
	end if
	response.redirect "Special.asp"
end if%>