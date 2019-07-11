<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
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
	
	if request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="check" or (request.cookies(Forcast_SN)("key")="selfreg" and jingyong=0) or request.cookies(Forcast_SN)("key")="typemaster" then
		NewsID=Request.QueryString("NewsID")
		set rs=server.CreateObject ("ADODB.RecordSet")
		rs.Source="select * from "& db_Type_Table &" order by typeorder"
		rs.Open rs.source,conn,1,1
		%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 修改文章类别</title>
</head>
<body topmargin="0">
<!--#include file="Admin_Top.asp"-->
<table border="1" width="100%" align=center cellpadding="0" cellspacing="0" style="border-collapse: collapse"  bordercolor="<%=border%>">
<form method="POST" action="NewsClassEdit2.asp?NewsID=<%=NewsID%>">
<tr align="center" >
<td colspan="2" height="25" width="100%"  class="TDtop">
	<div align="center">┊ <B>修改文章类别（修改ID号为<%=NewsID%>的文章的类别）</B> ┊</div>
</td>
</tr>
<tr>
<td colspan="2" align="center" height="160" bgcolor="#FFFFFF">请选择文章总栏：
	<select size="1" name="typeid">
	<%if rs.EOF then %>
<option value=0>暂无任何类别</option>
	<%else
		while not rs.EOF
		master=rs("typemaster")
		if master<>"" then
		tmaster=split(master,"|")
		for i=0 to ubound(tmaster)
			if request.cookies(Forcast_SN)("username")=trim(tmaster(i)) then
				typemaster=true
				exit for
			else
				typemaster=false
			end if
		next
	end if
	if (request.cookies(Forcast_SN)("key")="typemaster" and typemaster=true) or request.cookies(Forcast_SN)("key")="super" then%>
<option value='<%=rs("typeid")%>'><%=trim(rs("typeName"))%></option>
	<%end if
	rs.MoveNext
	wend
end if
%>	
</select>
</td>
</tr>
<tr>
<td colspan="2" width="100%" align="center" height="25" class="TDtop">
	<input type="button" value="返  回" onclick="javascript:history.go(-1)" class="unnamed5"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp;
	<input type="submit" value="下一步" name="B1"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
</td>
</tr>
</form>
</table>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
		<%
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
	end if
end if%>
