<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<%
response.buffer=true
select case request("action")
case "hidden"
	call hidden()
case "online"
	call online()
case else
	Errmsg=Errmsg+"<br>"+"<li>请指定正确的参数。"
	founderr=true
end select

call nav()
call head_var(1)
if founderr then
	call dvbbs_error()
else
	response.redirect Request.ServerVariables("HTTP_REFERER")
end if
call footer()

sub hidden()
if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>请登陆后进行操作。"
	founderr=true
	exit sub
end if

conn.execute("update online set userhidden=1 where userid="&userid)
dim usercookies
usercookies=request.cookies("aspsky")("usercookies")
if isnull(usercookies) or usercookies="" then usercookies="0"
select case usercookies
case "0"
	Response.Cookies("aspsky")("usercookies") = usercookies
case 1
   	Response.Cookies("aspsky").Expires=Date+1
	Response.Cookies("aspsky")("usercookies") = usercookies
case 2
	Response.Cookies("aspsky").Expires=Date+31
	Response.Cookies("aspsky")("usercookies") = usercookies
case 3
	Response.Cookies("aspsky").Expires=Date+365
	Response.Cookies("aspsky")("usercookies") = usercookies
end select
Response.Cookies("aspsky")("userhidden") = 1
Response.Cookies("aspsky").path=cookiepath
end sub

sub online()
if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>请登陆后进行操作。"
	founderr=true
	exit sub
end if

conn.execute("update online set userhidden=2 where userid="&userid)
dim usercookies
usercookies=request.cookies("aspsky")("usercookies")
if isnull(usercookies) or usercookies="" then usercookies="0"
select case usercookies
case "0"
	Response.Cookies("aspsky")("usercookies") = usercookies
case 1
   	Response.Cookies("aspsky").Expires=Date+1
	Response.Cookies("aspsky")("usercookies") = usercookies
case 2
	Response.Cookies("aspsky").Expires=Date+31
	Response.Cookies("aspsky")("usercookies") = usercookies
case 3
	Response.Cookies("aspsky").Expires=Date+365
	Response.Cookies("aspsky")("usercookies") = usercookies
end select
Response.Cookies("aspsky")("userhidden") = 2
Response.Cookies("aspsky").path=cookiepath
end sub
%>
