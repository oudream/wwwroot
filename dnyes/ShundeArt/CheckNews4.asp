<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	newsID=trim(request.QueryString("newsid"))
	set rs=server.createobject("adodb.recordset")
	sql="Select checkked,title from "& db_News_Table &" where newsID="&newsid
	rs.open sql,conn,1,3
	rs("checkked")="1" 
	rs.update
	dim title
	title=rs("title")
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "CheckNews.asp"
end if%>