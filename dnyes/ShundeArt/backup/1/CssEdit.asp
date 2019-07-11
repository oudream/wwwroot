<!--#include file="conn.asp" -->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp" -->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->

<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("KEY")="super") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	if css="1" or request.cookies(Forcast_SN)("purview")="99999" then
		%>
<html><head>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - CSS 编辑</title>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="site.css">
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<table border="1" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="<%=border%>" align=center>
<tr>
	<td height="30"> 
		<table border="1" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="<%=border%>">
		<tr class="TDtop">
			<td height=25 class="TDtop" align="center">┊ <B>文章字体模板编辑</B> ┊ </td>
		</tr>
		<tr bgcolor=ffffff> 
			<td valign=top> 
		<%
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		if request("save")="" then
			Set objCountFile = objFSO.OpenTextFile(Server.MapPath("news.css"),1,True)
			If Not objCountFile.AtEndOfStream Then fdata = objCountFile.ReadAll
		else
			fdata=request("fdata")
			Set objCountFile=objFSO.CreateTextFile(Server.MapPath("news.css"),True)
			objCountFile.Write fdata
		end if
		objCountFile.Close
		Set objCountFile=Nothing
		Set objFSO = Nothing
		%>
				<table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="<%=border%>">
				<form method=post>
				<tr> 
					<td width="97%" height="23" align="center">
						<table width="80%" border="0" cellspacing="0" cellpadding="0">
						<tr> 
							<td height="25" align="center">注意：文件将被保存在你安装目录下的<font color="#FF0000">news.css</font>里，如果您的空间不支持<font color="#FF0000">FSO</font>，请直接编辑该文件！</td>
						</tr>
						<tr> 
							<td height="25" align="center"><font color="#FF0000">如果您对CSS的属性不了解，请不要随意编辑。</font></td>
						</tr>
						</table>
					</td>
				</tr>
				<tr> 
					<td width="97%" align="center"> 
						<textarea name="fdata" cols="100" rows="20" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"><%=fdata%></textarea>
						<br>&nbsp;
					</td>
				</tr>
				<tr>
					<td height=25 class="TDtop" align="center"> 
						<input type="reset" name="Reset" value="重    置" style="font-family: 宋体; font-size: 9pt" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
						<input type="submit" name="save" value="提交修改" style="font-family: 宋体; font-size: 9pt" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
					</td>
				</tr>
				</form>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
	<%else
			Show_Err("对不起，该功能已经被超级系统管理员关闭，您没有权限操作！")
			response.end
	end if
end if%>