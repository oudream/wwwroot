<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp"-->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("�Բ������ĺ�̨����Ȩ�޲���������")
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
		Show_Message("<p align=center><font color=red>�û�ID ["&request("id")&"] ɾ���ɹ�!<br><br>"&freetime&"���Ӻ󷵻���ҳ!</font>")
		response.end
	else
		Show_Err("�Բ�������Ȩ�����û���ɾ����")
		response.end
	end if
end if%>