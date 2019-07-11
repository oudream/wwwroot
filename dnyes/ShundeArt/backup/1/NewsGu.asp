<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->

<%
if not(request.cookies(Forcast_SN)("KEY")="super" or request.cookies(Forcast_SN)("key")="bigmaster" or request.cookies(Forcast_SN)("key")="check" or request.cookies(Forcast_SN)("key")="typemaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	newsID=trim(request.QueryString("newsid"))
	set rs=server.createobject("adodb.recordset")
	sql="Select * from "& db_News_Table &" where newsID="&newsid
	rs.open sql,conn,1,3
	smallclassid=rs("smallclassid")
	bigclassid=rs("bigclassid")
	if rs("istop")=0 then
		rs("istop")="1"
	else
		rs("istop")="0"
	end if
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	
	if smallclassid<>"" then
		Response.Redirect "ListNews.asp?SmallClassID="&smallclassid
	else
		Response.Redirect "SmallNO.asp?BigClassID="&bigclassid
	end if
end if
%>