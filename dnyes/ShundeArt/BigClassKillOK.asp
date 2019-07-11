<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(Forcast_SN)("KEY")="super" or request.cookies(Forcast_SN)("KEY")="typemaster") then
	Show_Err("对不起，您的后台管理权限不够！")
	response.end
else
	delbigclass=request.form("delbigclass")
	
	if delbigclass="1" then
		BigClassid=request("BigClassid")
		dim bigclassname,typeid,typename
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_BigClass_Table &" where bigclassID="&bigclassID
		rs.open sql,conn,3,3
		bigclassname=rs("bigclassname")
		typeid=rs("typeid")
		rs.close
		set rs=nothing
		set rs=server.createobject("adodb.recordset")
		sql="select * from type where typeID="&typeID
		rs.open sql,conn,3,3
		typename=rs("typename")
		rs.close
		set rs=nothing
		button_value=trim(Request.Form("alert_button"))
		if button_value="是" then
			conn.execute("delete from news where bigclassid=" &bigclassid)
			conn.execute("delete from smallclass where bigclassid=" &bigclassid)
			conn.execute("delete from "& db_BigClass_Table &" where bigclassid=" &bigclassid)
			conn.close
			set conn=nothing
			response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=BigClassManage.asp?typeid="&typeid&""">"
			Show_Message("<p align=center><font color=red>恭喜您!您选择的大类已经被删除!<br><br>"&freetime&"秒钟后返回上页!</font>")
			response.end
		else
			response.redirect "BigClassManage.asp?typeid="&typeid
		end if
	else
		Show_Err("您无权删除此项目！<br><br><a href='javascript:history.back()'>返回</a>")
		response.end
	end if
end if
%>