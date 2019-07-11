<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(Forcast_SN)("KEY")="super" or request.cookies(Forcast_SN)("KEY")="bigmaster" or request.cookies(Forcast_SN)("KEY")="typemaster") THEN
	Show_Err("对不起，您的后台管理权限不够！")
	response.end
else
	delsmallclass=request.form("delsmallclass")
	if delsmallclass="1" then	
		dim smallclassname
		SmallClassid=request("SmallClassid")
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_SmallClass_Table &" where smallclassid="&SmallClassid
		rs.open sql,conn,1,1
		smallclassname=rs("smallclassname")
		bigclassid=rs("bigclassid")
		rs.close
		set rs=nothing
	
		button_value=trim(Request.Form("alert_button"))
		if button_value="是" then
			conn.execute("delete from "& db_News_Table &" where smallclassid=" &smallclassid)
			conn.execute("delete from "& db_SmallClass_Table &" where smallclassid=" &smallclassid)
			conn.close
			set conn=nothing
			response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=SmallClassManage.asp?bigclassid="&bigclassid&""">"
			Show_Message("<p align=center><font color=red>恭喜您!您删除小类“"&smallclassname&"”成功!!<br><br>"&freetime&"秒钟后返回上页!</font>")
			response.end
		else
			response.redirect "SmallClassManage.asp?bigclassid="&bigclassid
		end if
	else
		Show_Err("您无权删除此项目！<br><br><a href='javascript:history.back()'>返回</a>")
		response.end
	end if
end if%>