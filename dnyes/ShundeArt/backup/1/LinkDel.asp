<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="ChkURL.asp"-->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster" or request.cookies(Forcast_SN)("ManageKEY")="check" or request.cookies(Forcast_SN)("ManageKEY")="typemaster") then
	Show_Err("�Բ������Ĺ���Ȩ�޲�����")
	response.end
else
	if linkmana="1" or (request.cookies(Forcast_SN)("ManagePurview")="99999" and request.cookies(Forcast_SN)("purview")="99999") then
		ID=request("ID")
		if ID="" then
			Show_Err("��������Ҫɾ����ID��<br><br><a href='javascript:history.back()'>��ȥ����</a>")
			response.end
		end if
		set rs=server.CreateObject ("ADODB.RecordSet")
		rs.Source="select * from "& db_link_Table &" where ID=" & ID
		rs.Open rs.source,conn,1,1
		if rs.EOF then
			Show_Err("�Ƿ�ID����ȷ��ID�Ƿ���ڣ�<br><br><a href='javascript:history.back()'>��ȥ����</a>")
			Response.End
		end if
		%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - ɾ����������</title>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<table border="1" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="<%=border%>" align=center>
<form action=LinkDel2.asp method=post>
<tr class="TDtop">
	<td colspan=1 width="100%" height="25" align="center">�� <strong>ɾ �� �� վ �� �� ȷ ��</strong> ��</td>
</tr>
<input type="hidden" name="ID" value="<%=ID%>">
<tr>
	<td width="100%" align=center bgcolor="#FFFFFF">
		<p>&nbsp;</p>
		<p><font color=red>ɾ����վ���ӣ�[ <font color="#000066"> <%=ID%> </font>�� ]��<br><BR>
		��ɾ�����޷��ָ�����</font></p>
		<p>&nbsp;</p>
	</td>
</tr>
<tr class="TDtop">
	<td width="100%"align=center height="25" >
		<input type=submit value="  ��  " name="alert_button" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp;&nbsp;&nbsp;&nbsp;
		<input type=submit value="  ��  " name="alert_button" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
<tr>
	<td width="100%" align=center bgcolor="#FFFFFF"></td>
</tr>
</form>
</table>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
		<%rs.close 
		set rs=nothing
		conn.close
		set conn=nothing
	else
			Show_Err("�Բ��𣬸ù����Ѿ�������ϵͳ����Ա�رգ���û��Ȩ�޲�����")
		response.end
	end if
end if%>