<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<%
if request("action")="quit" then
	session("user_adminlevel")=0
	response.cookies(Forcast_SN)("UserName")=""
	response.cookies(Forcast_SN)("passwd")=""
	response.cookies(Forcast_SN)("UserEmail")=""
	response.cookies(Forcast_SN)("KEY")=""
	response.cookies(Forcast_SN)("purview")=""
	response.cookies(Forcast_SN)("fullname")=""
	response.cookies(Forcast_SN)("reglevel")=""
	response.cookies(Forcast_SN)("sex")=""
	response.cookies(Forcast_SN)("UserLoginTimes")=""
	if UserTableType = "Dvbbs" then	'是否整合论坛
		Response.Write "<iframe width='0' height='0' frameborder=no scrolling=no src='"& BbsPath &"logout.asp'></iframe>"
	end if

	if Request.cookies(Forcast_SN)("ManageUserName")<>"" then
		response.cookies(Forcast_SN)("ManageUserName")="" 
		response.cookies(Forcast_SN)("ManagePasswd")=""
		response.cookies(Forcast_SN)("ManageUserEmail")=""
		response.cookies(Forcast_SN)("ManageKEY")=""
		response.cookies(Forcast_SN)("ManagePurview")=""
		response.cookies(Forcast_SN)("ManageFullName")=""
		response.cookies(Forcast_SN)("ManageReglevel")=""
		response.cookies(Forcast_SN)("ManageSex")=""
		response.cookies(Forcast_SN)("ManageUserLoginTimes")=""
	end if
	response.redirect "index.asp"
end if %>
<HTML>
<HEAD>
<TITLE>用户退出登陆 </TITLE>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="site.css" rel=stylesheet type=text/css>
</HEAD>
<body bgcolor='#c1d2eb'>
<h1>
<p align="center">&nbsp;</p>
<p align="center">&nbsp;</p>
<p align="center">&nbsp;</p>
<p align="center"><b>确认退出网站登陆?</b></p></h1>
<p align="center">&nbsp;</p><h1>
<p align="center"><a href="Admin_Exit.asp?action=quit">退出</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="javascript:history.go(-1)">返回</a></p></h1>
</BODY>
</HTML>
