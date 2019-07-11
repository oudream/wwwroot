<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="md5.asp"-->
<!--#include file="ChkURL.asp"-->
<%
on error resume next
dim rs
UserName1=checkstr(request.form("UserName"))
passwd1=md5(trim(request.form("passwd")))
verifycode=cstr(trim(request.form("verifycode")))

dim ViewUrl
ViewUrl=request.cookies(Forcast_SN)("ViewUrl")
if ViewUrl="" then
	ViewUrl="index.asp"
end if

if cstr(session("verifycode"))<>verifycode then
	session("verifycode")=""
	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url="&Cstr(Request.ServerVariables("HTTP_REFERER"))&""">"
	Show_Message("<p align=center><font color=red>验证码出错，可能是您在某一页面停留太久了，请返回后刷新页面。</font><br><br>"&freetime&"秒钟后<a href=login.asp>返回</a>!")
	response.end
else
	session("verifycode")=""
	set rs=server.createobject("adodb.recordset")
	sql="select * from " & db_User_Table & " where " & db_User_Name & "='"&username1&"'"
	rs.open sql,ConnUser,1,3
	if passwd1<>rs(db_User_Password) or UserName1<>rs(db_User_Name)then
		response.write "<meta http-equiv=""refresh"" content="""&freetime&";url="&Cstr(Request.ServerVariables("HTTP_REFERER"))&""">"
		Show_Message("<p align=center><font color=red>用户名或密码出错，请返回检查.</font><br><br>"&freetime&"秒钟后<a href=login.asp>返回</a>!")
		response.end
	else
		rs(db_User_LastLoginIP)=Request.ServerVariables("REMOTE_ADDR")
		rs(db_User_LastLoginTime)=Now()
		rs(db_User_LoginTimes)=rs(db_User_LoginTimes)+1
		rs.update
		
		session("user_adminlevel")=1
		response.cookies(Forcast_SN)("UserName")=RS(db_User_Name)
		response.cookies(Forcast_SN)("passwd")=rs(db_User_Password)
		response.cookies(Forcast_SN)("UserEmail")=RS(db_User_Email)
		response.cookies(Forcast_SN)("KEY")=rs("OSKEY")
		response.cookies(Forcast_SN)("purview")=rs("purview")
		response.cookies(Forcast_SN)("fullname")=rs("fullname")
		response.cookies(Forcast_SN)("reglevel")=rs("reglevel")
		response.cookies(Forcast_SN)("sex")=rs(db_User_Sex)
		response.cookies(Forcast_SN)("UserLoginTimes")=rs(db_User_LoginTimes)
		rs.close
		set rs=nothing
	
		Conn.execute("delete from "& db_UploadPic_Table &" where username='"&request.cookies(Forcast_SN)("username")&"'")
	
		if UserTableType = "Dvbbs" then		'是否整合论坛
			%>
			<!--#include file="other_login.asp"-->
		<%else
			response.write "<meta http-equiv=""refresh"" content="""&freetime&";url="& ViewUrl &""">"
			Show_Message("<p align=center><font color=red>验证登录成功!</font><br><br>"&freetime&"秒钟后<a href="& ViewUrl &">返回</a>!")
			response.end
		end if
	end if
end if%>