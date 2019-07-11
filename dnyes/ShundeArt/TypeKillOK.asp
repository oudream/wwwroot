<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="char.inc"-->
<!--#include file="config.asp"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->
<%
IF request.cookies(Forcast_SN)("KEY")<>"super" THEN
	Show_Err("对不起，您的后台管理权限不够！")
	response.end
else
	typeID = Request("typeID")
	dim typename
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_Type_Table &" where typeID="&typeID
	rs.open sql,conn,3,3
	typename=rs("typename")
	rs.close
	set rs=nothing
	
	button_value=trim(Request.Form("alert_button"))
	if button_value="是" then
		conn.execute("delete from "& db_News_Table &" where typeID=" &typeID)
		conn.execute("delete from "& db_SmallClass_Table &" where typeID=" &typeID)
		conn.execute("delete from "& db_BigClass_Table &" where typeID=" &typeID)
		conn.execute("delete from "& db_Type_Table &" where typeID=" &typeID)
		conn.close
		set conn=nothing
		response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=TypeManage.asp"">"
		Show_Message("<p align=center><font color=red>恭喜您!您选择的总栏已经被删除!<br><br>"&freetime&"秒钟后返回上页!</font>")
		response.end
	else
		response.redirect "TypeManage.asp"
	end if
end if
%>