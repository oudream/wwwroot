<!--#include file="conn.asp"-->

<!--#include file="inc/const.asp"-->

<%
	dim activeuser
	membername=request.cookies("aspsky")("username")
	if session("userid")<>"" then
	activeuser="delete from online where id="&session("userid")
	Conn.Execute activeuser
	end if
	if membername<>"" then
	activeuser="delete from online where username='"&membername&"'"
	Conn.Execute activeuser
	end if
	Response.Cookies("aspsky").path=cookiepath
	Response.Cookies("aspsky")("username")=""
	Response.Cookies("aspsky")("password")=""
	Response.Cookies("aspsky")("userclass")=""
	Response.Cookies("aspsky")("userid")=""
	Response.Cookies("aspsky")("userhidden")=""
	Response.Cookies("aspsky")("usercookies")=""
	session("userid")=""
	conn.close
	set conn=nothing
	response.redirect("index.asp")
%>
