<%@ Language=VBScript %>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("�Բ������ĺ�̨����Ȩ�޲���������")
	response.end
else
	ReViewID=Request.Form("ReViewID")
	button_value=trim(Request.Form("alert_button"))
	if button_value="ȷ��" then
		conn.execute("delete from "& db_Review_Table &" where ReViewID=" & ReViewID)
	else
		Response.Redirect "ReViewManage.asp"
	end if

	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=ReViewManage.asp"">"
	Show_Message("���� ID =["& ReViewID &"]ɾ���ɹ���!<br><br>"&freetime&"���Ӻ󷵻���ҳ!</font>")
	response.end
end if
%>