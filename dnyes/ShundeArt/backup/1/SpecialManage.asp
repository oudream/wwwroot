<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="Config.ASP"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="Function.asp"-->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("�Բ������ĺ�̨����Ȩ�޲���������")
	response.end
else
	%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - ר�����</title>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<table width="100%" height="166" border="1" cellpadding="3" cellspacing="0" bordercolor="#C0C0C0" id="AutoNumber1" style="border-collapse: collapse">
<tr class="TDtop">
	<td  colspan="4" width="100%" height=25 align="center" valign="middle">��<strong>ר �� �� ��</strong>��</td>
</tr>
		<%
	Set rs6 = Server.CreateObject("ADODB.Recordset")
	sql6 ="SELECT  * From "& db_Special_Table &" order by SpecialID"
	RS6.open sql6,Conn,3,3
	%>
<tr align="center"  height=20>
	<td width="10%" >ID��</td>
	<td width="30%" >ר������</td>
	<td width="30%" >ר��ע��</td>
	<td width="30%" >ִ&nbsp;&nbsp;&nbsp;��</td>
</tr>
	<%do while not rs6.eof%>
<tr height=20>
	<td width="10%"  align="center" bgcolor="#FFFFFF"><%=rs6("SpecialID")%>��</td>
	<td width="30%"  align="center" bgcolor="#FFFFFF">
		<span style="CURSOR: hand" title="<%if rs6("specialzs")<>"" then%><%=nohtml(rs6("specialzs"))%><%else%>��<%end if%>"><%=nohtml(rs6("specialName"))%></span>
		<%if request.cookies(Forcast_SN)("ManageKEY")="super" then%>
		<font color=red>(<%=rs6("specialmaster")%>)</font>
		<%end if%>
	</td>
	<td width="30%" > &nbsp;<%=nohtml(rs6("Specialzs"))%>��</td>
	<td width="30%"  align="center" bgcolor="#FFFFFF">
		<%if (specialgl="1" and request.cookies(Forcast_SN)("ManageKEY")="bigmaster") or (shspecialgl="1" and request.cookies(Forcast_SN)("ManageKEY")="check") or request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or rs6("specialmaster")=request.cookies(Forcast_SN)("ManageUserName") then%>
		<a href="Specialedit.asp?SpecialID=<%=rs6("SpecialID")%>" onMouseOver="window.status='�༭ר�⡰<%=nohtml(rs6("SpecialName"))%>��������';return true;" onMouseOut="window.status='';return true;" title='�༭ר�⡰<%=nohtml(rs6("SpecialName"))%>��������'>�༭</a>
		<a href="Specialkill.asp?SpecialID=<%=rs6("SpecialID")%>" onMouseOver="window.status='ɾ��ר�⡰<%=nohtml(rs6("SpecialName"))%>��';return true;" onMouseOut="window.status='';return true;" title='ɾ��ר�⡰<%=nohtml(rs6("SpecialName"))%>��'>ɾ��</a>
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
		����ר�⣺<textarea name="Specialname" cols="20" rows="3" class="text" style="font-family: ����; font-size: 9pt" title="����������Ҫ��ӵ�ר������" onMouseOver="window.status='����������Ҫ��ӵ�ר������';return true;" onMouseOut="window.status='';return true;"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"></textarea>
		ר��ע�ͣ�<textarea name="specialzs" cols="40" rows="3" class="text" style="font-family: ����; font-size: 9pt" title="����������Ҫ��ӵ�ר��ע��" onMouseOver="window.status='����������Ҫ��ӵ�ר��ע��';return true;" onMouseOut="window.status='';return true;"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"></textarea>
		<input type="hidden" name="Specialmaster" value="<%=request.cookies(Forcast_SN)("ManageUserName")%>">
		<input type="submit" name="Submit" value="���" class=button onMouseOver="window.status='�������ť������ר��';return true;" onMouseOut="window.status='';return true;" title='�������ť������ר��' style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<input type="reset" value="��д" name="B1" class=button onMouseOver="window.status='�������ť�������ר��';return true;" onMouseOut="window.status='';return true;" title='�������ť�������ר��' style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
	</form>
</tr>
	<%end if%>
</table>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
<%end if%>