<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="Config.ASP"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<!--#include file="Function.asp"-->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	dim SpecialID
	SpecialID = Request("SpecialID")
	
	if SpecialID="" then
		Show_Err("未指定参数<br><br>－－－<a href='javascript:history.back()'>回去重来</a>－－－")
		response.end
	else
		if not IsNumeric(SpecialID) then
			Show_Err("非法参数<br><br>－－－<a href='javascript:history.back()'>回去重来</a>－－－")
			response.end
		else
		end if
	end if
	
	set rs=server.CreateObject("ADODB.RecordSet")
	rs.Source="select * from "& db_Special_Table &" where SpecialID=" & SpecialID
	rs.Open rs.Source,conn,1,1
	if rs("specialmaster")=request.cookies(Forcast_SN)("username") or request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("KEY")="check" or request.cookies(Forcast_SN)("KEY")="bigmaster" then
		%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 修改专题</title>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<div align="center">
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<form method="POST" action="SpecialEditok.asp">
<tr> 
	<td width="100%" height="25"  class="TDtop" ><p align="center" >┊ <strong>专题修改</strong> ┊</p></td>
</tr>
<tr align="center"> 
	<td width="100%"  bgcolor="#FFFFFF"> 
		<table border="0" cellspacing="1" width="90%" height="80">
			<tr> 
				<td align="center" height="60">
					<input type="hidden" name="SpecialID" value=<%=request("SpecialID")%>>
					<input type="hidden" name="Name" value='<%=nohtml(rs("SpecialName"))%>'>
					专题名称：<input type="text" name="SpecialName" size="20" maxlength="20" value='<%=nohtml(rs("SpecialName"))%>' style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"><br><br>
					专题注释：<input type="text" name="specialzs" size="20" class="text" value='<%=nohtml(rs("Specialzs"))%>' style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
				</td>
			</tr>
		</table>
	</td>
</tr>
<tr align="center"> 
	<td width="100%" class="TDtop1" height="60"> 
		<input type="submit" value=" 修 改 " name="cmdok" class=button  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp;
		<input type="reset" value=" 复 原 " name="cmdcancel" class=button style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp; 
		<input type="button" name="ok" value=" 放 弃 " onClick="javascript:history.go(-1)" class=button style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
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
		Show_Err("因为这不是您添加的专题，所以您无权编辑该专题！<br><br>－－－<a href='javascript:history.back()'>返回</a>－－－")
		response.end
	end if
end if%>