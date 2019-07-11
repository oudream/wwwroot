<!--#INCLUDE FILE="Conn.asp" -->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp" -->
<%
IF request.cookies(Forcast_SN)("ManageKEY")<>"super" then
	Show_Err("对不起，您的后台管理权限不够！")
	response.end
else
	dim jingyong
	dim ID
	ID=request("ID")
	jingyong=trim(request("jingyong"))
	
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_User_Table &" where "& db_User_ID &"="&ID
	rs.open sql,ConnUser,3,3
	rs("jingyong")=jingyong
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=usermanage.asp"">"
	Show_Message("<p align=center><font color=red>恭喜您!该会员状态设置成功!<br><br>"&freetime&"秒钟后返回上页!</font>")
	response.end
end if
%>