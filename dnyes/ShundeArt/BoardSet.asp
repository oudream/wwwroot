<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->

<%IF request.cookies(Forcast_SN)("KEY")<>"super" THEN
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	inuse=request.form("inuse")
	if inuse="" then
		Show_Err("对不起，请您选择公告项！<br><br><a href=history.go(-1)>返回</a>")
		response.end
	end if
	set rs=server.createobject("adodb.recordset")
	sql="Select inuse from "& db_board_Table &" where inuse=1"
	rs.open sql,conn,1,3
	if not rs.eof then
		do while not rs.eof
			rs("inuse")=0
		rs.movenext
		loop
	end if
	rs.close
	
	sql="Select * from "& db_board_Table &" where ID="&inuse
	rs.open sql,conn,1,3
	if not rs.EOF then
		do while not rs.EOF 
			rs("inuse")=1
			rs.MoveNext 
		loop
	end if
	rs.close
	set rs=nothing
	dim title
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_board_Table &" where inuse=1"
	rs.open sql,conn,1,1
	title=rs("title")
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.redirect "BoardManage.asp"
end if%>