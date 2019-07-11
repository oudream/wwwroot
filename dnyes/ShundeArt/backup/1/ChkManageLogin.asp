<!--#include file="conn.asp"-->
<!--#include file="char.inc"-->
<!--#include file="config.asp"-->
<!--#include file="md5.asp"-->

<%
on error resume next
dim rs
UserName1=CheckStr(request.form("UserName"))
passwd1=md5(trim(request.form("passwd")))
verifycode=cstr(trim(request.form("verifycode")))
dim ViewUrl1
ViewUrl1="Admin_login.asp"


if cstr(session("verifycode"))<>verifycode then
	Show_Err("验证码出错，可能是您在管理登录页面停留太久了，请返回后刷新页面。<br><br><a href='javascript:history.back()'>返回</a>")
	response.end
else
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_Manage_Table &" where ("& db_ManageUser_Name &"='"&username1&"' and adder='"&Request.cookies(Forcast_SN)("username")&"')"
	rs.open sql,Conn,1,3
	if passwd1<>rs(db_ManageUser_Password) or UserName1<>rs(db_ManageUser_Name)then
		Show_Err("用户名或密码出错，请返回检查<br><br><a href='javascript:history.back()'>返回重来</a>")
		response.end
	else
		rs(db_ManageUser_LastLoginIP)=Request.ServerVariables("REMOTE_ADDR")
		rs(db_ManageUser_LastLoginTime)=Now()
		rs(db_ManageUser_LoginTimes)=rs(db_ManageUser_LoginTimes)+1
		rs.update
		
		response.cookies(Forcast_SN)("ManageUserName")=RS(db_ManageUser_Name)
		response.cookies(Forcast_SN)("ManagePasswd")=rs(db_ManageUser_Password)
		response.cookies(Forcast_SN)("ManageUserEmail")=RS(db_ManageUser_Email)
		response.cookies(Forcast_SN)("ManageKEY")=rs("OSKEY")
		response.cookies(Forcast_SN)("ManagePurview")=rs("purview")
		response.cookies(Forcast_SN)("ManageFullName")=rs("fullname")
		response.cookies(Forcast_SN)("ManageReglevel")=rs("reglevel")
		response.cookies(Forcast_SN)("ManageSex")=rs(db_ManageUser_Sex)
		response.cookies(Forcast_SN)("ManageUserLoginTimes")=rs(db_ManageUser_LoginTimes)
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
	
		response.redirect "Admin_login.asp"
		response.end
	end if
end if
rs.close
set rs=nothing
Conn.close
set Conn=nothing
%>