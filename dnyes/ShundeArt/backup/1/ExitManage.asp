<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<%
if request("action")="quit" then
	response.cookies(Forcast_SN)("ManageUserName")="" 
	response.cookies(Forcast_SN)("ManagePasswd")=""
	response.cookies(Forcast_SN)("ManageUserEmail")=""
	response.cookies(Forcast_SN)("ManageKEY")=""
	response.cookies(Forcast_SN)("ManagePurview")=""
	response.cookies(Forcast_SN)("ManageFullName")=""
	response.cookies(Forcast_SN)("ManageReglevel")=""
	response.cookies(Forcast_SN)("ManageSex")=""
	response.cookies(Forcast_SN)("ManageUserLoginTimes")=""
	response.redirect "Admin_login.asp"
end if
Show_Message("<h1><p align='center'>&nbsp;</p><p align='center'><b>ȷ���˳���վ����?</b></p></h1><p align='center'>&nbsp;</p><h1><p align='center'><a href='ExitManage.asp?action=quit'>�˳�</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=javascript:history.go(-1)>����</a></p></h1>")
response.end
%>