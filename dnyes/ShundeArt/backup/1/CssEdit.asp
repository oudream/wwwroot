<!--#include file="conn.asp" -->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp" -->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->

<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("KEY")="super") then
	Show_Err("�Բ������ĺ�̨����Ȩ�޲���������")
	response.end
else
	if css="1" or request.cookies(Forcast_SN)("purview")="99999" then
		%>
<html><head>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - CSS �༭</title>
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
			<td height=25 class="TDtop" align="center">�� <B>��������ģ��༭</B> �� </td>
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
							<td height="25" align="center">ע�⣺�ļ������������㰲װĿ¼�µ�<font color="#FF0000">news.css</font>�������Ŀռ䲻֧��<font color="#FF0000">FSO</font>����ֱ�ӱ༭���ļ���</td>
						</tr>
						<tr> 
							<td height="25" align="center"><font color="#FF0000">�������CSS�����Բ��˽⣬�벻Ҫ����༭��</font></td>
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
						<input type="reset" name="Reset" value="��    ��" style="font-family: ����; font-size: 9pt" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
						<input type="submit" name="save" value="�ύ�޸�" style="font-family: ����; font-size: 9pt" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
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
			Show_Err("�Բ��𣬸ù����Ѿ�������ϵͳ����Ա�رգ���û��Ȩ�޲�����")
			response.end
	end if
end if%>