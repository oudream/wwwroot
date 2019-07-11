<!--#include file="SafeRequest.asp"-->
<!--#include file="Char.inc"-->
<%
if not(Request.cookies(Forcast_SN)("KEY")="super" or Request.cookies(Forcast_SN)("KEY")="check" or Request.cookies(Forcast_SN)("KEY")="typemaster") THEN
	Show_Err("对不起，您无权登录后台进行管理！")
	response.end
else
	if ((Request.cookies(Forcast_SN)("ManageUserName")="") or (Request.cookies(Forcast_SN)("ManagePasswd")="") or (Request.cookies(Forcast_SN)("ManageKEY")="")) then
		if IsDebug="1" then
			Show_Err("cookies(ManageUserName)(ManagePasswd)(ManageKEY)中有为空值。<br><br><a href='javascript:history.back()'>返回</a>")
		else
			response.redirect "Admin_login.asp?action=Manage.asp"
		end if
		response.end
	else
		IF not(Request.cookies(Forcast_SN)("ManageKEY")="super" or Request.cookies(Forcast_SN)("ManageKEY")="check" or Request.cookies(Forcast_SN)("ManageKEY")="typemaster" or Request.cookies(Forcast_SN)("ManageKEY")="bigmaster" or Request.cookies(Forcast_SN)("ManageKEY")="smallmaster" ) THEN
			if IsDebug="1" then
					Show_Err("cookies(KEY)值不是[ManageSuper][ManageCheck][ManageTypemaster][ManageBigmaster][ManageSmallmaster]中的一个。<br><br><a href='javascript:history.back()'>返回</a>")
			else
				response.redirect "Admin_login.asp?action=Manage.asp"
			end if
			response.end
		else
			UserName=CheckStr(trim(Request.cookies(Forcast_SN)("UserName")))
			M_UserName=CheckStr(trim(Request.cookies(Forcast_SN)("ManageUserName")))
			M_PassWD=trim(Request.cookies(Forcast_SN)("ManagePasswd"))
			M_Key=trim(request.cookies(Forcast_SN)("ManageKEY"))
			set urs=server.createobject("adodb.recordset")
			sql="select * from "& db_Manage_Table &" where ("& db_ManageUser_Name &"='"& M_UserName &"' and adder='"& UserName &"')"
			urs.open sql,Conn,1,3
			if urs.bof or urs.eof then
				if IsDebug="1" then
					Show_Err("以cookies(username)及cookies(ManageUserName)的返回值，查找不到管理用户表中的相对应记录。<br><br><a href='javascript:history.back()'>返回</a>")
				else
					response.redirect "Admin_login.asp?action=Manage.asp"
				end if
				response.end
			else
				if M_PassWD<>urs(db_ManageUser_Password) then
					if IsDebug="1" then
						Show_Err("cookies(ManagePasswd)的返回值与管理用户表中的相对应记录的Password值不符。<br><br><a href='javascript:history.back()'>返回</a>")
					else
						response.redirect "Admin_login.asp?action=Manage.asp"
					end if
					response.end
				else
					if M_key<>urs("OSKEY") then
						if IsDebug="1" then
							Show_Err("cookies(ManageKEY)的返回值与管理用户表中的相对应记录的OSKEY值不符。<br><br><a href='javascript:history.back()'>返回</a>")
						else
							response.redirect "Admin_login.asp?action=Manage.asp"
						end if
						response.end
					end if
				end if
				urs.close
				set urs=nothing
			end if
		end if
	end if
end if%>