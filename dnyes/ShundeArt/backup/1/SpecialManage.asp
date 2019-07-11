<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="Config.ASP"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="Function.asp"-->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 专题管理</title>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<table width="100%" height="166" border="1" cellpadding="3" cellspacing="0" bordercolor="#C0C0C0" id="AutoNumber1" style="border-collapse: collapse">
<tr class="TDtop">
	<td  colspan="4" width="100%" height=25 align="center" valign="middle">┊<strong>专 题 管 理</strong>┊</td>
</tr>
		<%
	Set rs6 = Server.CreateObject("ADODB.Recordset")
	sql6 ="SELECT  * From "& db_Special_Table &" order by SpecialID"
	RS6.open sql6,Conn,3,3
	%>
<tr align="center"  height=20>
	<td width="10%" >ID号</td>
	<td width="30%" >专题名称</td>
	<td width="30%" >专题注释</td>
	<td width="30%" >执&nbsp;&nbsp;&nbsp;行</td>
</tr>
	<%do while not rs6.eof%>
<tr height=20>
	<td width="10%"  align="center" bgcolor="#FFFFFF"><%=rs6("SpecialID")%>　</td>
	<td width="30%"  align="center" bgcolor="#FFFFFF">
		<span style="CURSOR: hand" title="<%if rs6("specialzs")<>"" then%><%=nohtml(rs6("specialzs"))%><%else%>无<%end if%>"><%=nohtml(rs6("specialName"))%></span>
		<%if request.cookies(Forcast_SN)("ManageKEY")="super" then%>
		<font color=red>(<%=rs6("specialmaster")%>)</font>
		<%end if%>
	</td>
	<td width="30%" > &nbsp;<%=nohtml(rs6("Specialzs"))%>　</td>
	<td width="30%"  align="center" bgcolor="#FFFFFF">
		<%if (specialgl="1" and request.cookies(Forcast_SN)("ManageKEY")="bigmaster") or (shspecialgl="1" and request.cookies(Forcast_SN)("ManageKEY")="check") or request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or rs6("specialmaster")=request.cookies(Forcast_SN)("ManageUserName") then%>
		<a href="Specialedit.asp?SpecialID=<%=rs6("SpecialID")%>" onMouseOver="window.status='编辑专题“<%=nohtml(rs6("SpecialName"))%>”的属性';return true;" onMouseOut="window.status='';return true;" title='编辑专题“<%=nohtml(rs6("SpecialName"))%>”的属性'>编辑</a>
		<a href="Specialkill.asp?SpecialID=<%=rs6("SpecialID")%>" onMouseOver="window.status='删除专题“<%=nohtml(rs6("SpecialName"))%>”';return true;" onMouseOut="window.status='';return true;" title='删除专题“<%=nohtml(rs6("SpecialName"))%>”'>删除</a>
		<%end if%>
	</td>
</tr>
		<%rs6.MoveNext
	Loop
	rs6.close
	set rs6=nothing
	conn.close
	set conn=nothing
	if (specialgl="1" and request.cookies(Forcast_SN)("Managekey")="bigmaster") or (shspecialgl="1" and request.cookies(Forcast_SN)("ManageKEY")="check") or request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" then%>
<tr valign="middle">
	<form method="post" action="SpecialAdd.asp" name="type">
	<td width="100%" height="60" colspan="4" align="center">
		增加专题：<textarea name="Specialname" cols="20" rows="3" class="text" style="font-family: 宋体; font-size: 9pt" title="在这里输入要添加的专题名称" onMouseOver="window.status='在这里输入要添加的专题名称';return true;" onMouseOut="window.status='';return true;"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"></textarea>
		专题注释：<textarea name="specialzs" cols="40" rows="3" class="text" style="font-family: 宋体; font-size: 9pt" title="在这里输入要添加的专题注释" onMouseOver="window.status='在这里输入要添加的专题注释';return true;" onMouseOut="window.status='';return true;"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"></textarea>
		<input type="hidden" name="Specialmaster" value="<%=request.cookies(Forcast_SN)("ManageUserName")%>">
		<input type="submit" name="Submit" value="添加" class=button onMouseOver="window.status='按这个按钮添加这个专题';return true;" onMouseOut="window.status='';return true;" title='按这个按钮添加这个专题' style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<input type="reset" value="重写" name="B1" class=button onMouseOver="window.status='按这个按钮重新添加专题';return true;" onMouseOut="window.status='';return true;" title='按这个按钮重新添加专题' style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
	</form>
</tr>
	<%end if%>
</table>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
<%end if%>