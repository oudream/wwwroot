<!--#INCLUDE FILE="Conn.asp" -->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->
<%
IF request.cookies(Forcast_SN)("KEY")="" THEN
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	if request("action")="update" then
		dim smallclassorder,smallclassma,smallclassview,smallclassname,bigclassid,smallclasszs
		for i=1 to request.form("smallclassid").count
			smallclassid=CheckStr(request.form("smallclassid")(i))
			smallclassorder=CheckStr(request.form("smallclassorder")(i))
			smallclassma=CheckStr(request.form("smallclassma")(i))
			smallclassview=CheckStr(request.form("smallclassview")(i))
			smallclassname=CheckStr(request.form("smallclassname")(i))
			bigclassid=CheckStr(request.form("bigclassid")(i))
			smallclasszs=CheckStr(request.form("smallclasszs")(i))
			if CheckStr(request.form("smallclassorder")(i))="" then
				Show_Err("请填写小类排序！<br><br><a href=history.go(-1)>返回</a>")
				response.end
			end if
			if CheckStr(request.form("smallclassma")(i))="" then
				smallclassma="无"
			else
				master=split(CheckStr(request.form("smallclassma")(i)),"|")
			 	for k=0 to ubound(master)
					set rs=server.createobject("adodb.recordset")
					sql="Select * from "& db_User_Table &" where oskey='smallmaster' and  "& db_User_Name &"='"&trim(master(k))&"'"
					rs.open sql,ConnUser,1,3
					if trim(master(k))<>"无" then
						if rs.eof then
							Show_Err("小类管理员中无“& trim(master(k)) &”用户！请重新选择该小类的管理员！<br><br><a href=history.go(-1)>返回</a>")
							Response.End
						end if
					else
						smallclassma="无"
					end if
					rs.close
					set rs=nothing
				next
			end if
			conn.execute("update "& db_smallclass_Table &" set smallclassorder="&smallclassorder&",smallclassma='"&smallclassma&"',smallclassview="&smallclassview&",smallclassname='"&smallclassname&"',bigclassid="&bigclassid&",smallclasszs='"&smallclasszs&"' where smallclassid="&smallclassid)
			Set nrs = Server.CreateObject("ADODB.Recordset")
			sqln="select * from "& db_News_Table &" where smallclassid="&smallclassid
			nrs.open sqln,conn,3,3
			while not nrs.EOF
				nrs("bigclassid")=bigclassid
				nrs.MoveNext
			wend
			nrs.close
			set nrs=nothing
		next
	end if

	if request("action")="add" then
		function changechr(str) 
			changechr=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," "," ") 
			changechr=replace(changechr,"'","&acute;")
			changechr=replace(changechr,mid(" "" ",2,1),"&quot;")
		end function
					
		smallclasszs=request.form("smallclasszs")
		smallclassname=changechr(trim(request("smallclassname")))
		smallclassorder=request.form("smallclassorder")
		smallclassma=request.form("smallclassma")
		smallclassview=request.form("smallclassview")
		bigclassid=request.form("bigclassid")
		typeid=request.form("typeid")
		
		If smallclassname="" Then
			Show_Err("小类名称不能为空！请<a href=javascript:history.go(-1)>返回重新填写</a>！")
			response.end
		end if
		If smallclassorder="" Then
			Show_Err("小类排序不能为空！请<a href=javascript:history.go(-1)>返回重新填写</a>！")
			response.end
		end if
		
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql="select * from "& db_SmallClass_Table &""
		rs.open sql,conn,3,3
	 	rs.addnew
		rs("smallclassname")=smallclassname
		rs("smallclasszs")=smallclasszs
		rs("bigclassid")=bigclassid
		rs("typeid")=typeid
		if smallclassorder="" then
			rs("smallclassorder")=0
		else
			rs("smallclassorder")=smallclassorder
		end if
		rs("smallclassview")=smallclassview
		if smallclassma="" then
			rs("smallclassma")="无"
		else
			rs("smallclassma")=smallclassma
		end if
		rs.update
		bigclassid=rs("bigclassid")
		rs.close
		set rs=nothing
	end if
		
	conn.close
	set conn=nothing
	response.write "<meta http-equiv=""refresh"" content="""&freetime&";url=SmallClassManage.asp?bigclassid="&bigclassid&""">"
	Show_Message("<p align=center><font color=red>恭喜您!设置成功!<br><br>"&freetime&"秒钟后返回上页!</font>")
	response.end
end if
%>