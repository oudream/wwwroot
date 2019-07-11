<!--#INCLUDE FILE="Conn.asp" -->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->
<%
IF request.cookies(Forcast_SN)("KEY")<>"super" THEN
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	if request("action")="update" then
		dim bigclassorder,bigclassmaster,bigclassview,bigclassname,typeid,bigclasszs
		for i=1 to request.form("bigclassid").count
			bigclassid=CheckStr(request.form("bigclassid")(i))
			bigclassorder=CheckStr(request.form("bigclassorder")(i))
			bigclassmaster=CheckStr(request.form("bigclassmaster")(i))
			bigclassview=CheckStr(request.form("bigclassview")(i))
			bigclassname=CheckStr(request.form("bigclassname")(i))
			bigclasszs=CheckStr(request.form("bigclasszs")(i))
			typeid=CheckStr(request.form("typeid")(i))
			if CheckStr(request.form("bigclassorder")(i))="" then
				Show_Err("请填写大类排序！<br><br><a href=history.go(-1)>返回</a>")
				response.end
			end if
			if replace(request.form("bigclassmaster")(i),"'","")="" then
				bigclassmaster="无"
			else
				master=split(CheckStr(request.form("bigclassmaster")(i)),"|")
		 		for k=0 to ubound(master)
					set rs=server.createobject("adodb.recordset")
					sql="Select * from "& db_User_Table &" where oskey='bigmaster' and "& db_User_Name &"='"&trim(master(k))&"'"
					rs.open sql,ConnUser,1,3
					if trim(master(k))<>"无" then
						if rs.eof then
							Show_Err("大类管理员中无“& trim(master(k)) &”用户！请重新选择该大类的管理员！<br><br><a href=history.go(-1)>返回</a>")
							Response.End
						end if
					else
						bigclassmaster="无"
					end if
					rs.close
					set rs=nothing
		 		next
			end if
		
			conn.execute("update "& db_BigClass_Table &" set bigclassorder="&bigclassorder&",bigclassmaster='"&bigclassmaster&"',bigclassview="&bigclassview&",bigclassname='"&bigclassname&"',typeid="&typeid&",bigclasszs='"&bigclasszs&"' where bigclassid="&bigclassid)
		    Set nrs = Server.CreateObject("ADODB.Recordset")
		    sqln="select * from news where bigclassid="&bigclassid
		    nrs.open sqln,conn,3,3
		    while not nrs.EOF
		    nrs("typeid")=typeid
		    nrs.MoveNext
		    wend
		    nrs.close
		    set nrs=nothing
		
		next
	end if
	
	if request("action")="add" then
	bigclasszs=request.form("bigclasszs")
	bigclassname=trim(request("bigclassname"))
	bigclassorder=request.form("bigclassorder")
	bigclassmaster=request.form("bigclassmaster")
	bigclassview=request.form("bigclassview")
	typeid=request.form("typeid")
	
	If bigclassorder="" Then
		Show_Err("大类排序不能为空！请<a href=javascript:history.go(-1)>返回重新填写</a>！")
		response.end
	end if
	If bigclassname="" Then
		Show_Err("大类名称不能为空！请<a href=javascript:history.go(-1)>返回重新填写</a>！")
		response.end
	end if
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from "& db_BigClass_Table &""
	rs.open sql,conn,3,3
	rs.addnew
	rs("bigclassname")=bigclassname
	rs("bigclasszs")=bigclasszs
	rs("typeid")=typeid
	if bigclassorder="" then
		rs("bigclassorder")=0
	else
		rs("bigclassorder")=bigclassorder
	end if
	rs("bigclassview")=bigclassview
	if bigclassmaster="" then
		rs("bigclassmaster")="无"
	else
		rs("bigclassmaster")=bigclassmaster
	end if
	rs.update
	typeid=rs("typeid")
	rs.close
	set rs=nothing
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from type where typeid="&typeid
	rs.open sql,conn,3,3
	dim typename
	typename=rs("typename")
	rs.close
	set rs=nothing
	end if
	
	conn.close
	set conn=nothing
	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=BigClassManage.asp?typeid="&typeid&""">"
	Show_Message("<p align=center><font color=red>恭喜您!设置成功!<br><br>"&freetime&"秒钟后返回上页!</font>")
	response.end
end if
%>