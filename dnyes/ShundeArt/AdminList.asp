<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="md5.asp"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp"-->
<!--#include file="ChkURL.asp"-->

<%
IF not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
else
	dim sql
	dim rst
	set rst=server.createobject("adodb.recordset")
	sql="select * from "& db_Manage_Table &" where "& db_ManageUser_ID &"="&request("id")
	rst.open sql,Conn,1,1
	%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - <%=rst("fullname")%> 的详细资料</title>
</head>
<body topmargin="0">
<!--#include file="Admin_Top.asp"-->
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr>
	<td colspan="2">
		<div align="right">编号：<%=rst(db_ManageUser_Id)%>&nbsp;&nbsp;
	<%if rst(db_ManageUser_RegDate)<>"" then%>
			注册日期：<%=rst(db_ManageUser_RegDate)%> 
	<%end if%>
		</div>
	</td>
</tr>
<tr> 
	<td width="614"> 
	<%if rst("fullname")<>"" then%>
		&nbsp;&nbsp;&nbsp;&nbsp;姓&nbsp;&nbsp;&nbsp; 名：<%=rst("fullname")%><br> <br> 
	<%end if%>
	<%if db_Sex_Select = "FeiTium" then%>
		&nbsp;&nbsp;&nbsp;&nbsp;性&nbsp;&nbsp;&nbsp;&nbsp;别：<%=rst(db_ManageUser_sex)%> 
	<%else%>
		<%if db_Sex_Select = "Number" then%>
		&nbsp;&nbsp;&nbsp;&nbsp;性&nbsp;&nbsp;&nbsp;&nbsp;别：<%if rst(db_ManageUser_Sex)="0" then%>女<%else%>男<%end if%>
		<%end if%>
	<%end if%>
	<br> <br> 
	<%if db_Birthday_Select = "FeiTium" then%>
		&nbsp;&nbsp;&nbsp;&nbsp;生&nbsp;&nbsp;&nbsp; 日：<%=rst(db_ManageUser_Birthyear)%>年<%=rst(db_ManageUser_Birthmonth)%>月<%=rst(db_ManageUser_Birthday)%>日 
	<%else%>
		<%if db_Birthday_Select = "Text" then%>
		&nbsp;&nbsp;&nbsp;&nbsp;生&nbsp;&nbsp;&nbsp; 日：<%=year(rst(db_ManageUser_birthday))%>年<%=month(rst(db_ManageUser_birthday))%>月<%=day(rst(db_ManageUser_birthday))%>日 
		<%end if%>
	<%end if%>
	<br> <br> 
	<%if rst(db_ManageUser_Email)<>"" then%>
		&nbsp;&nbsp;&nbsp;&nbsp;电子邮件：<a href=mailto:<%=rst(db_ManageUser_Email)%>><%=rst(db_ManageUser_Email)%></a><br> 
	<br> 
	<%end if%>
	<%if rst("tel")<>"" then%>
		&nbsp;&nbsp;&nbsp;&nbsp;联系电话：<%=rst("tel")%><br> <br> 
	<%end if%>
	&nbsp;&nbsp;&nbsp;&nbsp;权&nbsp;&nbsp;&nbsp; 限：
	<%oskey1=rst("oskey")
	select case oskey1
		case "super"
			if rst("purview")="99999" then
				response.write "超级管理员"
			else
				response.write "系统管理员"
			end if
		case "typemaster" response.write "总栏管理员"
		case "bigmaster" response.write "大类管理员"
		case "smallmaster" response.write "小类管理员"
		case "selfreg" response.write "注册用户"
		case "checkf" response.write "审核管理员"
	end select%>
		<br><br>&nbsp;&nbsp;&nbsp;&nbsp;工作单位：<%=rst("depname")%>
		<br><br>&nbsp;&nbsp;&nbsp;&nbsp;所在部门：<%=rst("deptype")%>
		<br><br>&nbsp;&nbsp;&nbsp;&nbsp;登录次数：<%=rst(db_ManageUser_LoginTimes)%>
		<br><br>&nbsp;&nbsp;&nbsp;&nbsp;文&nbsp;章&nbsp;数：<%=rst("number")%>
		<br><br>&nbsp;&nbsp;&nbsp;&nbsp;最近地址：<%=rst(db_ManageUser_LastLoginIP)%>
		<br><br>&nbsp;&nbsp;&nbsp;&nbsp;最近登录：<%=rst(db_ManageUser_LastLoginTime)%>
		<br><br>&nbsp;&nbsp;&nbsp;&nbsp;自我介绍：<%=rst("content")%>
	</td>
	<td width="137"><div align="center"> 
		<%if rst(db_User_Face) <> "" then%>
			<%if UserTableType = "FeiTium" then%>
		         <img src="<%=rst(db_User_Face)%>" border="0" width="120">
			<%else%>
				 <%if UserTableType = "Dvbbs" then%>
					<img src="<%=Bbspath%><%=rst(db_User_Face)%>" border="0" width="<%=rst(db_User_FaceWidth)%>" height="<%=rst(db_User_FaceHeight)%>"> 
					<%''显示用户头像，加'bbs/'前缀路径,使图文系统直接显示定向到论坛头像%>
				 <%end if%>
			<%end if%>
		<%else%>
			<img src="images/nopic.gif" border="0"> 
		<%end if%>
		<br>
		<br>
		用户头像</div>
    </td>
</tr>
<tr> 
	<td colspan="2">
		<div align="right">[<a href="AdminEdit.asp?id=<%=rst(db_ManageUser_Id)%>">修改</a>] 
		<%if trim(session("username"))<>trim(rst(db_ManageUser_Name)) then%>
            <% if trim(request.cookies(Forcast_SN)("username"))=trim(rst(db_ManageUser_Name)) then %>
			    <%response.write "[----]"%>
			<%else%>
			    [<a href="AdminDel.asp?id=<%=rst(db_ManageUser_Id)%>&amp;name=del">删除</a>] 
		    <%end if%>
		<%end if%>
			[<a href="javascript:history.go(-1)">返回</a>]
		</div>
	</td>
</tr>
</table>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
<%
END IF
rst.close
set rst=nothing
conn.close
set conn=nothing
%>