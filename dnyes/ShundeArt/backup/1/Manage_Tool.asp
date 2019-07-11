<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp" -->
<%
IF request.cookies(Forcast_SN)("ManageKEY")<>"super" then
	Show_Err("对不起，您的后台管理权限不够！")
	response.end
else
	response.buffer=true
	Response.Expires=0
	%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 部门管理</title>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<center>
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr class="TDtop"> 
	<td  colspan="4" width="100%" height=25 align="center" valign="middle" height="24">
		┊ <strong>运行管理工具</strong> ┊
	</td>
</tr>
<tr> 
	<td colspan="4" align="center" width="100%" height=25 bgcolor="#FFFFFF"></td>
</tr>
<form method="post" action="Manage_Tool.asp" name="type">
<tr> 
	<td align="center" colspan="4" width="100%"  height="34">
		程序名称：<select name="tool_file_name" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	<%toolfilename=request("tool_file_name")

	dim theFolder
	dim theFiles
	dim cpath,lpath
	set fsoBrowse=CreateObject("Scripting.FileSystemObject")

	lpath=left(Request.ServerVariables("PATH_INFO"),instrrev(Request.ServerVariables("PATH_INFO"),"/"))
	cpath=Server.MapPath(lpath)
	
	if fsoBrowse.FolderExists(cpath) then
		Set theFolder=fsoBrowse.GetFolder(cpath)
		Set theFiles=theFolder.files

		for each x in theFiles
			if UCase(left(x.Name,6))="TOOLS_" then
				if toolfilename<>"" then
					if x.Name=toolfilename then
						response.write "<option value='"& x.Name &"' selected>"& x.Name &"</option>"
					else
						response.write "<option value='"& x.Name &"'>"& x.Name &"</option>"
					end if
				else
					response.write "<option value='"& x.Name &"'>"& x.Name &"</option>"
				end if
			end if
		next
	end if
	%>
		</select>&nbsp;&nbsp;
		<input type="submit" name="Submit" value="运行" title="运行选择的工具程序" onMouseOver="window.status='运行选择的工具程序';return true;" onMouseOut="window.status='';return true;" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp;&nbsp;
	</td>
</tr>
<tr> 
	<td colspan="4" align="center" width="100%" height="25" bgcolor="#FFFFFF"></td>
</tr>
<tr> 
	<td colspan="4" align="center" width="100%" height="250" bgcolor="#FFFFFF">
	<%if toolfilename<>"" then%>
		<iframe style="top:2px" ID="RunTools" src='<%=toolfilename%>' frameborder=0 width="100%" height="100%"></iframe></
	<%end if%>
	</td>
</tr>
</form>
</table>
</center>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
<%end if%>