<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp" -->

<%IF not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster" or request.cookies(Forcast_SN)("ManageKEY")="check" or request.cookies(Forcast_SN)("ManageKEY")="typemaster") THEN
	Show_Err("�Բ������Ĺ���Ȩ�޲�����")
	response.end
else
	aaas=(request.cookies(Forcast_SN)("ManageKEY")="super")
	if votemana="1" or (request.cookies(Forcast_SN)("ManagePurview")="99999" and request.cookies(Forcast_SN)("purview")="99999") then
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_Vote_Table &" order by id desc"
		rs.open sql,conn,1,1
		%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - ͶƱ����</title>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<form method="POST" action="VoteSet.asp">
<tr class="TDtop">
	<td width="100%" height="25" colspan=5 align=center>�� <strong>ͶƱ����</strong> ��</td>
</tr>
<tr class="TDtop1" height=20>
	<td width="10%" align="center">ѡ��</td>
	<td width="10%" align="center">ID</td>
	<td width="60%" align="center">����</td>
	<td width="10%" align="center">�޸�</td>
	<td width="10%" align="center">ɾ��</td>
</tr>
		<%do while not rs.eof%>
<tr height=20>
	<td width="10%" align="center">
		<input type="radio" value=<%=rs("ID")%><%if rs("IsChecked")=1 then%> checked<%end if%> name="Checked">
	</td>
	<td width="10%" align="center"><%=rs("ID")%>��</td>
	<td width="60%"><%=rs("Title")%>��</td>
	<td width="10%" align="center">
		<input onclick="javascript:window.open('VoteModify.asp?id=<%=rs("ID")%>','_self','')" type="button" value="�޸�" name="button1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
	<td width="10%" align="center">
		<input onclick="javascript:window.open('VoteDel.asp?id=<%=rs("ID")%>','_self','')" type="button" value="ɾ��" name="button2" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
			<%rs.movenext
		loop
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		%>
<tr>
	<td colspan=5 align=center>
		<input type="submit" value="ѡ��ͶƱ��" name="submit" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<input onclick="javascript:window.open('VoteAdd.asp','_self','')" type="button" value="�����ͶƱ" name="button" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
</form>
</table>
</div>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
	<%else
		Show_Err("�Բ��𣬸ù����Ѿ�������ϵͳ����Ա�رգ���û��Ȩ�޲�����")
		response.end
	end if
end if%>