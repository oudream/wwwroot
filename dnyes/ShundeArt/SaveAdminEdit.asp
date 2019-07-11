<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="md5.asp"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp"-->
<!--#include file="ChkURL.asp"-->

<%
IF not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or  request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	dim sql
	dim sql11
	dim rs
	dim rs11
	dim id
	dim passwd,passwd1
	dim username
	dim oskey
	dim depid
	dim shenhe
	dim fullname
	
	username=htmlencode(request.form("username"))
	fullname=htmlencode(request.form("fullname"))
	passwd=htmlencode(request.form("passwd"))
	passwd1=md5(trim(request.form("passwd")))
	oskey=htmlencode(request.form("oskey"))
	purview=(request.form("purview"))
	id=request("id")
	depid=request.form("depid")
	shenhe=request.form("shenhe")
	
	dim depname
	dim deptype
	set rs11=server.createobject("adodb.recordset")
	sql11="select * from "& db_Dep_Table &" where id="&depid
	rs11.open sql11,conn,3,3
	depname=rs11("depname")
	deptype=rs11("deptype")
	rs11.close
	
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_Manage_Table &" where "& db_ManageUser_ID &"="&id
	rs.open sql,Conn,3,3
	rs(db_ManageUser_Name)=username
	rs("fullname")=fullname
	rs("purview")=purview
	
	if passwd<>rs(db_ManageUser_Password) then
		rs(db_ManageUser_Password)=passwd1
	end if
	
	if oskey<>"" then
		rs("oskey")=oskey
	end if
	rs("depid")=depid
	rs("depname")=depname
	rs("deptype")=deptype
	if shehe<>"" then
		rs("shenhe")=shenhe
	end if
	rs.update
	rs.close
	
	conn.close
	set conn=nothing
	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=AdminManage.asp"">"
	Show_Message("<p align=center><font color=red>恭喜您!该用户资料已经修改成功!<br><br>"&freetime&"秒钟后返回上页!</font>")
	response.end
end if
%>