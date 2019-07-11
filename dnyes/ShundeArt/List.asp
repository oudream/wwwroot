<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp" -->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	dim sql
	dim rst
	set rst=server.createobject("adodb.recordset")
	sql="select * from "& db_User_Table &" where "&Db_User_ID&"="&request("id")
	rst.open sql,ConnUser,1,1
	if UserTableType = "Dvbbs" then
		photowidth=rst(db_User_FaceWidth)		''取论坛设定的图片宽度
		photoheight=rst(db_User_FaceHeight)		''取论坛设定的图片高度
	end if
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
		<div align="right">
			编号：<%=rst(db_User_Id)%>&nbsp;&nbsp; 
	<%if rst(db_User_RegDate)<>"" then%>
			注册日期：<%=rst(db_User_RegDate)%> 
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
		&nbsp;&nbsp;&nbsp;&nbsp;性&nbsp;&nbsp;&nbsp;&nbsp;别：<%=rst(db_User_sex)%> 
	<%else%>
		<%if db_Sex_Select = "Number" then%>
		&nbsp;&nbsp;&nbsp;&nbsp;性&nbsp;&nbsp;&nbsp;&nbsp;别： 
			<%if rst(db_User_Sex)="0" then%>
		女
			<%else%>
		男
			<%end if%>
		<%end if%>
	<%end if%>
	<br> <br> 
	<%if db_Birthday_Select = "FeiTium" then%>
		&nbsp;&nbsp;&nbsp;&nbsp;生&nbsp;&nbsp;&nbsp; 日：<%=rst(db_User_Birthyear)%>年<%=rst(db_User_Birthmonth)%>月<%=rst(db_User_Birthday)%>日 
	<%else%>
		<%if db_Birthday_Select = "Text" then%>
		&nbsp;&nbsp;&nbsp;&nbsp;生&nbsp;&nbsp;&nbsp; 日：<%=year(rst(db_User_birthday))%>年<%=month(rst(db_User_birthday))%>月<%=day(rst(db_User_birthday))%>日 
		<%end if%>
	<%end if%>
	<br> <br> 
	<%if rst(db_User_Email)<>"" then%>
		&nbsp;&nbsp;&nbsp;&nbsp;电子邮件：<a href=mailto:<%=rst(db_User_Email)%>><%=rst(db_User_Email)%></a><br> 
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
				response.write "超级用户"
			else
				response.write "系统用户"
			end if
		case "typemaster" response.write "总栏用户"
		case "bigmaster" response.write "大类用户"
		case "smallmaster" response.write "小类用户"
		case "selfreg" response.write "注册用户"
		case "checkf" response.write "审核用户"
	end select%>
		<br> <br> &nbsp;&nbsp;&nbsp;&nbsp;工作单位：<%=rst("depname")%><br> <br>
		&nbsp;&nbsp;&nbsp;&nbsp;所在部门：<%=rst("deptype")%><br> <br>&nbsp;&nbsp;&nbsp;&nbsp;登录次数：<%=rst(db_User_LoginTimes)%><br> 
		<br> &nbsp;&nbsp;&nbsp;&nbsp;文&nbsp;章&nbsp;数：<%=rst("number")%><br> <br> 
		&nbsp;&nbsp;&nbsp;&nbsp;最近地址：<%=rst(db_User_LastLoginIP)%><br> <br> &nbsp;&nbsp;&nbsp;&nbsp;最近登录：<%=rst(db_User_LastLoginTime)%><br> 
		<br> &nbsp;&nbsp;&nbsp;&nbsp;自我介绍：<%=rst("content")%> </td>
		<td width="137"><div align="center"> 
		<%if rst(db_User_Face) <> "" then
			if UserTableType = "FeiTium" then%>
		<img src="<%=rst(db_User_Face)%>" border="0" width="120"> 
			<%else
				if UserTableType = "Dvbbs" then%>
					<img src="<%=Bbspath%><%=rst(db_User_Face)%>" border="0" width="<%=rst(db_User_FaceWidth)%>" height="<%=rst(db_User_FaceHeight)%>"> 
					<%''显示用户头像，加'bbs/'前缀路径,使图文系统直接显示定向到论坛头像%>
				<%end if
			end if
		else%>
			<img src="images/nopic.gif" border="0"> 
		<%end if%>
		<br>
		<br>
		用户头像</div></td>
</tr>
<tr> 
	<td colspan="2">
		<div align="right">[<a href="UserEdit.asp?id=<%=rst(db_User_Id)%>">修改</a>] 
		<%if trim(session("username"))<>trim(rst(db_User_Name)) then%>
            <% if trim(request.cookies(Forcast_SN)("username"))=trim(rst(db_User_Name)) then %>
			    <%response.write "[----]"%>
			<%else%>
			    [<a href="UserDel.asp?id=<%=rst(db_User_Id)%>&amp;name=del">删除</a>] 
		    <%end if%>
		<%end if%>
		[<a href="javascript:history.go(-1)">返回</a>]</div></td>
</tr>
</table>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
	<%
	rst.close
	set rst=nothing
	conn.close
	set conn=nothing
end if%>