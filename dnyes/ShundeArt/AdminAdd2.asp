<!--#include file=conn.asp -->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="md5.asp"-->
<!--#include file="ChkURL.asp"-->
<!--#include file="ChkManage.asp"-->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	usernamecookie=CheckStr(request.cookies(Forcast_SN)("UserName"))
	passwdcookie=CheckStr(trim(Request.cookies(Forcast_SN)("passwd")))
	KEYcookie=CheckStr(trim(request.cookies(Forcast_SN)("KEY")))
	
	dim rs,tsql
	dim rst
	dim UserName
	dim passwd
	dim adder
	dim depid
	dim oskey
	dim shenhe
	adder=request.form("adder")
	UserName=trim(request("UserName"))
	depid=request.form("depid")
	oskey=request.form("oskey")
	passwd=md5(trim(request("passwd")))
	purview=trim(request("purview"))
	
	if Instr(request("UserName"),"=")>0 or Instr(request("UserName"),"%")>0 or Instr(request("UserName"),chr(32))>0 or Instr(request("UserName"),"?")>0 or Instr(request("UserName"),"&")>0 or Instr(request("UserName"),";")>0 or Instr(request("UserName"),",")>0 or Instr(request("UserName"),"'")>0 or Instr(request("UserName"),",")>0 or Instr(request("UserName"),chr(34))>0 or Instr(request("UserName"),chr(9))>0 or Instr(request("UserName"),"")>0 or Instr(request("UserName"),"$")>0 then
		Show_Err("对不起，您输入的用户名中包含有非法字符。<br><br><a href='javascript:history.back()'>回去重来</a>")
		response.end
	end if
	
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_Manage_Table &""
	rs.open sql,Conn,3,3
	do while not rs.eof
		if rs(db_ManageUser_Name)=UserName then
			Show_Err("对不起，已经存在这个用户["& UserName &"]。<br><br><a href='javascript:history.back()'>回去重来</a>")
			response.end
		end if
		rs.movenext
	loop
	rs.close
	set rs=nothing
	
	dim depname
	dim deptype
	set rs11=server.createobject("adodb.recordset")
	sql11="select * from "& db_Dep_Table &" where id="&depid
	rs11.open sql11,conn,3,3
	depname=rs11("depname")
	deptype=rs11("deptype")
	rs11.close
	set rs11=nothing
	
	set rst=server.CreateObject("ADODB.RecordSet")
	rst.open "select * from "& db_Manage_Table,Conn,3,2
	rst.addnew
	rst("oskey")=oskey
	rst(db_ManageUser_Name)=Username
	rst(db_ManageUser_Password)=Passwd
	rst("purview")=purview
	rst("depid")=depid
	rst("fullname")=request.form("fullname")
	rst("depname")=depname
	rst("deptype")=deptype
	rst("adder")=adder
	rst("jingyong")=0
	rst("reglevel")=1
	rst(db_ManageUser_RegDate )=now()
	if shenhe="" then
		rst("shenhe")=1
	else
		rst("shenhe")=request.form("shenhe")
	end if
	rst.update
	oskey=rst("oskey")
	rst.close
	
	response.redirect "AdminManage.asp"
end if%>