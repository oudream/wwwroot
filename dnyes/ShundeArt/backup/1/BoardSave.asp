<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp" -->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster" or request.cookies(Forcast_SN)("ManageKEY")="check" or request.cookies(Forcast_SN)("ManageKEY")="typemaster") then
	Show_Err("�Բ������Ĺ���Ȩ�޲�����")
	response.end
else
	ID=Request.QueryString("ID")
	Title=Request.Form("Title")

	if Title="" then
		Show_Err("����д���±��⣡<br><br><a href='javascript:history.back()'>��ȥ����</a>")
		response.end
	end if

	if len(title)>100 then
		Show_Err("������������<br><br><a href='javascript:history.back()'>��ȥ����</a>")
		response.end
	end if

	Content=CheckStr(Request.Form("Content"))
	if Content="" then
		Show_Err("�����빫�����ݣ�<br><br><a href='javascript:history.back()'>��ȥ����</a>")
		response.end
	end if

	upload=CheckStr(Request.Form("upload"))
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_board_Table
	rs.open sql,conn,1,3
	rs.addnew
	rs("content")=content
	rs("upload")=upload
	rs("title")=title
	rs("dateandtime")=now()
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "boardManage.asp"
end if
%>