<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(Forcast_SN)("KEY")="super" or request.cookies(Forcast_SN)("KEY")="check" or request.cookies(Forcast_SN)("key")="typemaster" or request.cookies(Forcast_SN)("key")="bigmaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	newsid=trim(request.QueryString("newsid"))
	set rs=server.createobject("adodb.recordset")
	sql="Select * from "& db_News_Table &" where newsID="&newsid
	rs.open sql,conn,1,3
	smallclassid=rs("smallclassid")
	bigclassid=rs("bigclassid")
	
	if rs("checkked")=0 then
		rs("checkked")=1
	else
		rs("checkked")=0
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