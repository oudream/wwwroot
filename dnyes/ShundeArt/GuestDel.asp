<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkURL.asp"-->
<%
if request.cookies(Forcast_SN)("key")="super" then
	reviewid=cint(request("reviewid"))
	conn.execute "delete from "& db_Review_Table &" where reviewid="&reviewid
	conn.close
	set conn=nothing
	response.redirect "GuestBook.asp"
else
	Show_Err("����Ȩɾ�������ԣ�<br><br><a href='javascript:history.back()'>����</a>")
	response.end
end if
%>