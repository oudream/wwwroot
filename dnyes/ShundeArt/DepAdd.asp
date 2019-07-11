<!--#INCLUDE FILE="Conn.asp" -->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp" -->
<%
IF request.cookies(Forcast_SN)("ManageKEY")<>"super" then
	Show_Err("对不起，您的后台管理权限不够！")
	response.end
else
	function changechr(str) 
		changechr=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," "," ") 
		changechr=replace(changechr,"'","&acute;")
		changechr=replace(changechr,mid(" "" ",2,1),"&quot;")
	end function

	dim depname
	dim deptype
	deptype=changechr(trim(request.form("deptype")))
	depname=changechr(trim(request("depname")))
	
	If depname="" Then
		Show_Err("对不起，单位名称不能为空！<br><br><a href='javascript:history.back()'>返回</a>")
		response.end
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_dep_Table &""
	rs.open sql,conn,3,3
	do while not rs.eof
		if rs("depname")=depname and rs("deptype")=deptype then
			Show_Err("对不起，相同的单位部门已经存在！<br><br><a href='javascript:history.back()'>返回</a>")
			response.end
		end if
		rs.movenext
	loop

If deptype="" Then
		Show_Err("对不起，部门名称不能为空！<br><br><a href='javascript:history.back()'>返回</a>")
		response.end
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_dep_Table &""
	rs.open sql,conn,3,3
	do while not rs.eof
		
		rs.movenext
	loop

	rs.close
	set rs=nothing
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from "& db_dep_Table &""
	rs.open sql,conn,3,3
	rs.addnew
	rs("depname")=depname
	rs("deptype")=deptype
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.redirect "DepManage.asp"
end if%>