<!--#include file="config.asp"-->
<%
adsconn.open adsdata
dim adssql,adsrs,admname,admpass
admname=trim(request("username"))
admpass=trim(request("password"))
set adsrs=server.createobject("adodb.recordset")
asql="select * from admin where admin='"&admname&"'"
adsrs.open asql,adsconn,1,1
if adsrs.bof and adsrs.eof then
%>
无此管理员！
<%
else
if admpass=adsrs("passwd") then
session("masterlogin")="superadvertadmin"
session("admname")=adsrs("admin")
session("aid")=adsrs("aid")
Response.Redirect "index.asp"
adsrs.close
set adsrs=nothing
adsconn.close
set adsconn=nothing
else
%>
密码错误！
<%
end if
end if
%>