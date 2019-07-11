<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	dim jingyong
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_User_Table &" where  "& db_User_Name &"='"&request.cookies(Forcast_SN)("username")&"'"
	rs.open sql,ConnUser,1,3
	jingyong=rs("jingyong")
	rs.close
	set rs=nothing
	
	if request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="check" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster" or request.cookies(Forcast_SN)("ManageKEY")="smallmaster" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or (request.cookies(Forcast_SN)("ManageKEY")="selfreg" and jingyong=0) then
		NewsID=request("NewsID")
		if NewsID="" then
			Show_Err("文章ID号不能为空！<br><br>－－－<a href='javascript:history.back()'>回去重来</a>－－－")
			response.end
		end if
		set rs=server.CreateObject ("ADODB.RecordSet")
		rs.Source="select * from "& db_News_Table &" where NewsID=" & NewsID
		rs.Open rs.source,conn,1,1
		if rs.EOF then
				Show_Err("非法的文章ID号["& NewsID &"]，请确认NewsID不超出当前范围！<br><br>－－－<a href='javascript:history.back()'>回去重来</a>－－－")
				response.end
		end if
		username=rs("editor")
		typeid=rs("typeid")
		bigclassid=rs("bigclassid")
		smallclassid=rs("smallclassid")
		if smallclassid<>"" then
		set rs3=server.createobject("adodb.recordset")
		sql3="select * from "& db_SmallClass_Table &" where smallclassid="&smallclassid
		rs3.open sql3,conn,1,3
		master6=rs3("smallclassma")
		if master6<>"" then
		smallmaster6=split(master6,"|")
		 for i=0 to ubound(smallmaster6)
			if request.cookies(Forcast_SN)("username")=trim(smallmaster6(i)) then
			smallclassmaster6=true
			exit for
			else
			smallclassmaster6=false
			end if
		 next
		end if
		rs3.close
		set rs3=nothing
		end if
		set rs1=server.createobject("adodb.recordset")
		sql1="select * from "& db_BigClass_Table &" where bigclassid="&rs("bigclassid")
		rs1.open sql1,conn,1,3
		master5=rs1("bigclassmaster")
		if master5<>"" then
		bigmaster5=split(master5,"|")
		 for i=0 to ubound(bigmaster5)
			if request.cookies(Forcast_SN)("username")=trim(bigmaster5(i)) then
			bigclassmaster5=true
			exit for
			else
			bigclassmaster5=false
			end if
		 next
		end if
		set rs2=server.createobject("adodb.recordset")
		sql2="select * from "& db_Type_Table &" where typeid="&typeid
		rs2.open sql2,conn,1,3
		master4=rs2("typemaster")
		if master4<>"" then
		tmaster4=split(master4,"|")
		 for i=0 to ubound(tmaster4)
			if request.cookies(Forcast_SN)("username")=trim(tmaster4(i)) then
			typemaster4=true
			exit for
			else
			typemaster4=false
			end if
		 next
		end if
		rs2.close
		set rs2=nothing
		rs1.close
		set rs1=nothing

		if request.cookies(Forcast_SN)("key")="super" or (request.cookies(Forcast_SN)("key")="bigmaster" and bigclassmaster5=true) or (request.cookies(Forcast_SN)("key")="typemaster" and typemaster4=true) or (request.cookies(Forcast_SN)("key")="check" and checkdel="1") or (request.cookies(Forcast_SN)("key")="smallmaster" and smallclassmaster6=true and request.cookies(Forcast_SN)("username")=username) or (request.cookies(Forcast_SN)("key")="selfreg" and jingyong=0 and request.cookies(Forcast_SN)("username")=username) then
			%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 删除文章</title>
</head>
<body topmargin="0"><!--#include file=Admin_Top.asp-->
<div align="center">
<table border="1" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="<%=border%>">
<form action=NewsDel2.asp method=post>
<tr> 
	<td width="100%" height="25"  class="TDtop" ><p align="center" >┊ <strong>删 除 文 章 确 认</strong> ┊</p></td>
</tr>
			<%set rs2=server.createobject("adodb.recordset")
			sql2="select * from "& db_User_Table &" where  "& db_User_Name &"='"&username&"'"
			rs2.open sql2,ConnUser,1,3
			if rs2.EOF then
				nouser=1
			end if
			rs2.close
			set rs2=nothing
			%>
<input type="hidden" name="NewsID" value="<%=NewsID%>">
<tr>
<td width="100%" align=center bgcolor="#FFFFFF">
<p>　</p>
<p><font color=red>删除文章：[ <font color="#000066"> <%=trim(rs("Title"))%>
</font>]？<br><%if rs("checkked")="1" then%>这篇文章已经通过审核，并且已经发表！<br>确实要删除吗？<%else%>这篇文章尚未通过审核！<br>确实要删除吗？<%end if%>
<BR>
（删除后将无法恢复！）</font></p>
<p><%if nouser=1 then%>该文章的作者已不在本系统中，是否还要删除？<input type="hidden" value="1" name="del"><%end if%></p>
</td>
</tr>
<tr> 
	<td width="100%" height="25"  class="TDtop" align=center>
		<input type="hidden" name="delnews" value="1">
		<input type=submit value="  是  " name="alert_button" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp;&nbsp;&nbsp;&nbsp;
		<input type=submit value="  否  " name="alert_button" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
</form>
</table>
</div>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
		<%rs.close
			set rs=nothing
			conn.close
			set conn=nothing
		else
			Show_Err("－－－您无权删除－－－")
			response.end
		end if
	end if
end if%>