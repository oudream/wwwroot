<%@ Language=VBScript %>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="function.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	dim jingyong,delnews
	delnews=request.form("delnews")
	if delnews="1" then
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_User_Table &" where  "& db_User_Name &"='"&request.cookies(Forcast_SN)("username")&"'"
		rs.open sql,ConnUser,1,3
		jingyong=rs("jingyong")
		rs.close
		set rs=nothing

		if request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="check" or request.cookies(Forcast_SN)("key")="bigmaster" or request.cookies(Forcast_SN)("key")="smallmaster" or (request.cookies(Forcast_SN)("key")="selfreg" and jingyong=0) or request.cookies(Forcast_SN)("key")="typemaster" then%>
			<%dim username,typeid,typename,bigclassid,bigclassname,smallclassid,smallclassname
			NewsID=Request.Form("NewsID")
			set rs=server.CreateObject("ADODB.RecordSet")
			rs.Source="select * from "& db_News_Table &" where NewsID=" & NewsID
			rs.Open rs.source,conn,1,1
			image=rs("image")
			username=rs("editor")
			dim title
			title=rs("title")
			newsrelated=rs("newsrelated")
			typeid=rs("typeid")
			bigclassid=rs("bigclassid")
			smallclassid=rs("smallclassid")
			rs.close
			set rs=nothing
			set rs=server.CreateObject("ADODB.RecordSet")
			rs.Source="select * from "& db_Type_Table &" where typeID=" & typeID
			rs.Open rs.source,conn,1,1
			typename=rs("typename")
			rs.close
			set rs=nothing
			set rs=server.CreateObject("ADODB.RecordSet")
			rs.Source="select * from "& db_BigClass_Table &" where bigclassID=" & bigclassID
			rs.Open rs.source,conn,1,1
			bigclassname=rs("bigclassname")
			rs.close
			set rs=nothing
			if smallclassid<>"" then
				set rs=server.CreateObject("ADODB.RecordSet")
				rs.Source="select * from "& db_SmallClass_Table &" where smallclassID=" & smallclassID
				rs.Open rs.source,conn,1,1
				smallclassname=rs("smallclassname")
				rs.close
				set rs=nothing
			end if
			button_value=trim(Request.Form("alert_button"))
			if button_value="是" then
				if request.Form("del")<>"1" then
					set rs2=server.createobject("adodb.recordset")
					sql2="select * from "& db_User_Table &" where  "& db_User_Name &"='"&username&"'"
					rs2.open sql2,ConnUser,1,3
					rs2("number")=rs2("number")-1
					rs2.update
					rs2.close
					set rs2=nothing
				end if
				conn.execute("delete from "& db_News_Table &" where NewsID=" & NewsID)
				conn.execute("delete from "& db_Review_Table &" where NewsID=" & NewsID)
				set rs2=server.createobject("adodb.recordset")
				sql2="select * from "& db_Attach_Table &" where NewsID=" & NewsID
				rs2.open sql2,conn,1,3
				do while not rs2.EOF
					Set fso=Server.CreateObject("Scripting.FileSystemObject")
					If fso.FileExists(Server.Mappath(""& FileUploadPath & rs2("filename")))=true Then
						fso.DeleteFile Server.MapPath(""& FileUploadPath & rs2("filename") )
					End If
							Set fso=Nothing
					rs2.movenext
				loop
				rs2.close
				set rs2=nothing
				conn.execute("delete from "& db_Attach_Table &" where NewsID=" & NewsID)
				if smallclassid<>"" then
	response.write "<p align=center><font color=red>恭喜您!您选择的总栏“"&typename&"”下大类“"&bigclassname&"”下的小类“"&smallclassname&"”的文章“"&title&"”已经被删除！<br>"&freetime&"秒钟后返回上页！</font>"
	else
	response.write "<p align=center><font color=red>恭喜您!您选择的总栏“"&typename&"”下大类“"&bigclassname&"”的文章“"&title&"”已经被删除！<br>"&freetime&"秒钟后返回上页！</font>"
	end if
	if smallclassid<>"" then
	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=listnews.asp?smallclassid="&smallclassid&""">"
	else
	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=smallno.asp?bigclassid="&bigclassid&""">"
	end if
	else
	response.write "<p align=center><font color=red>您没有执行删除操作！<br>"&freetime&"秒钟后返回上页！</font>"
	if smallclassid<>"" then
	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=listnews.asp?smallclassid="&smallclassid&""">"
	else
	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=smallno.asp?bigclassid="&bigclassid&""">"
	end if
	end if
	end if
	else
	response.write "<script>alert('您无权删除此项目！');history.go(-1);</Script>"
	response.end
	end if
end if
%>