<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster" or request.cookies(Forcast_SN)("ManageKEY")="check" or request.cookies(Forcast_SN)("ManageKEY")="typemaster") then
	Show_Err("�Բ������Ĺ���Ȩ�޲�����")
	response.end
else
	id=request.QueryString("id")
	set rs=server.createobject("adodb.recordset")
	sql="delete from "& db_board_Table &" where ID="&id
	rs.open sql,conn,1,3
	conn.close
	set conn=nothing
	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=BoardManage.asp"">"
	Show_Message("<p align=center><font color=red>��ϲ��!����ɾ���ɹ�!<br><br>"&freetime&"���Ӻ󷵻���ҳ!</font>")
	response.end
end if%>