<%@ Language=VBScript %>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp" -->
<%
if not((request.cookies(Forcast_SN)("ManageKEY")="super" and request.cookies(Forcast_SN)("ManagePurview")="99999") or (request.cookies(Forcast_SN)("KEY")="super" and request.cookies(Forcast_SN)("purview")="99999")) then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	if init="1" then
		if (request.form("action")="del") then
			if request.form("selectdel")<>"" then
				For i = 1 To Request.Form("selectdel").Count
					SelectTable=Request.Form("selectdel")(i)
					select case SelectTable
						'初始化部门表
						case db_Dep_Table conn.execute "delete from "& db_Dep_Table
						'初始化总栏表
						case db_Type_Table conn.execute "delete from "& db_Type_Table
						'初始化大栏表
						case db_BigClass_Table conn.execute "delete from "& db_BigClass_Table
						'初始化小栏表
						case db_SmallClass_Table conn.execute "delete from "& db_SmallClass_Table
						'初始化专题表
						case db_Special_Table conn.execute "delete from "& db_Special_Table
						'初始化文章表
						case db_News_Table conn.execute "delete from "& db_News_Table
						'初始化临时表
						case db_NewsFile_Table conn.execute "delete from "& db_NewsFile_Table
						'初始化上传临时表
						case db_UploadPic_Table conn.execute "delete from "& db_UploadPic_Table
						'初始化投票表
						case db_Vote_Table conn.execute "delete from "& db_Vote_Table
						'初始化友情链接表
						case db_Link_Table conn.execute "delete from "& db_Link_Table
						'初始化公告表
						case db_Board_Table conn.execute "delete from "& db_Board_Table
						'初始化附件表
						case db_Attach_Table conn.execute "delete from "& db_Attach_Table
						'初始化评论表
						case (db_Review_Table&"1") conn.execute "delete from "& db_Review_Table &" where NewsId>0"
						'初始化留言表
						case (db_Review_Table&"0") conn.execute "delete from "& db_Review_Table &" where NewsId<1"
						'初始化用户表
						case db_User_Table ConnUser.execute "delete from "& db_User_Table &" where purview<>99999"
						'初始化管理员表
						case db_Manage_Table conn.execute "delete from "& db_Manage_Table &" where purview<>99999"
					end select
				Next
				ConnUser.close
				set ConnUser=nothing
				conn.close
				set conn=nothing
				response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=New.asp"">"
				Show_Message("<p align=center><font color=red>系统表"& (request.form("selectdel")) &"初始化完成!<br><br>"&freetime&"秒钟后返回初始页!</font>")
				response.end
			else
				Show_Err("对不起，您并没有选择任何操作！<br><br>－－－<a href='javascript:history.back()'>回去重来</a>－－－")
				response.end
			end if
		else
			SpecialID=Request.Form("SpecialID")
			button_value=trim(Request.Form("alert_button"))
			if button_value="确定" then
				conn.execute("delete from "& db_BigClass_Table &"")
				conn.execute("delete from "& db_SmallClass_Table &"")
				conn.execute("delete from "& db_Special_Table &"")
				conn.execute("delete from "& db_News_Table &"")
				conn.execute("delete from "& db_Review_Table &"")
				conn.execute("delete from "& db_Vote_Table &"")
				conn.execute("delete from "& db_Link_Table &"")
				conn.execute("delete from "& db_Type_Table &"")
				conn.execute("delete from "& db_Dep_Table &"")
				conn.execute("delete from "& db_Board_Table &"")
				conn.execute("delete from "& db_Attach_Table &"")
				conn.execute("delete from "& db_NewsFile_Table &"")
				conn.execute("delete from "& db_UploadPic_Table &"")
				if UserTableType = "FeiTium" then		'不整合则
					ConnUser.execute("delete from "& db_User_Table &" where purview<>99999")
				end if
				response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=Admin_login.asp"">"
				Show_Message("<p align=center><font color=red>系统已经被初始化!<br><br>"&freetime&"秒钟后返回首页!</font>")
				response.end
			else
				Response.Redirect "Admin_login.asp"
			end if
		end if
	else
		Show_Err("对不起，该功能已经被超级系统管理员关闭，您没有权限操作！")
		response.end
	end if
end if%>