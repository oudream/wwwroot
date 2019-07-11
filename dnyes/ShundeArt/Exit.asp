<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<%
dim ViewUrl
ViewUrl=request.cookies(Forcast_SN)("ViewUrl")
if ViewUrl="" then
	ViewUrl="index.asp"
end if
if UserTableType = "Dvbbs" then		'是否整合论坛
	Response.Write "<iframe width='0' height='0' frameborder=no scrolling=no src='"& BbsPath &"logout.asp'></iframe>"
end if
response.cookies(Forcast_SN)("UserName")=""
response.cookies(Forcast_SN)("KEY")=""
response.cookies(Forcast_SN)("purview")=""
response.cookies(Forcast_SN)("fullname")=""
response.cookies(Forcast_SN)("passwd")=""
response.cookies(Forcast_SN)("sex")=""
response.cookies(Forcast_SN)("reglevel")=""
response.redirect ViewUrl
%>