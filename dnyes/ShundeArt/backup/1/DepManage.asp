<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp" -->
<!--#include file="Function.asp"-->
<%
IF request.cookies(Forcast_SN)("ManageKEY")<>"super" then
	Show_Err("�Բ������ĺ�̨����Ȩ�޲�����")
	response.end
else
	%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - ��λ���Ź���</title>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<center>
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr class="TDtop"> 
	<td  colspan="4" width="100%" height=25 align="center" valign="middle" height="24">
		�� <strong>��λ���Ź���</strong> ��
	</td>
</tr>
	<%dim deptype
	deptype=request("deptype")
	Set rs6 = Server.CreateObject("ADODB.Recordset")
	if deptype="" then
		sql6 ="select * from "& db_dep_Table &" order by ID"
	else
		sql6 ="select * from "& db_dep_Table &" where deptype="&deptype&" order by ID"
	end if
	rs6.open sql6,Conn,3,3
	%>
<tr align="center" class="TDtop1"> 
	<td width="3%" >ԭ��</td>
	<td width="14%" >��λ����</td>
	<td width="6%" height="25" >��������</td>
	<td width="10%" >��λ�༭</td>
</tr>
	<%do while not rs6.eof%>
<tr> 
	<td width="3%" align="center" bgcolor="#FFFFFF"><%=rs6("ID")%></td>
	<td width="14%" align="center" bgcolor="#FFFFFF"><%=nohtml(rs6("depName"))%></span></td>
	<td width="6%" height="25" align="center" bgcolor="#FFFFFF"> 
		<%deptype=nohtml(rs6("deptype"))
		response.write ""& deptype &""
		%>
	</td>
	<td width="10%" bgcolor="#FFFFFF" align="center">
		<a href="DepEdit.asp?ID=<%=rs6("ID")%>&depName=<%=nohtml(rs6("depName"))%>&deptype=<%=nohtml(rs6("deptype"))%>" onMouseOver="window.status='�༭��λ��<%=nohtml(rs6("depName"))%>��������';return true;" onMouseOut="window.status='';return true;" title='�༭��λ��<%=nohtml(rs6("depName"))%>��������'>�༭</a>&nbsp;
		<a href="DepKill.asp?ID=<%=rs6("ID")%>&depName=<%=nohtml(rs6("depName"))%>" onMouseOver="window.status='ɾ����λ��<%=nohtml(rs6("depName"))%>��';return true;" onMouseOut="window.status='';return true;" title='ɾ����λ��<%=nohtml(rs6("depName"))%>��'>ɾ��</a>
	</td>
</tr>
		<%
		rs6.MoveNext
	loop
	rs6.close
	set rs6=nothing
	%>
<tr> 
	<td colspan="4" height="40" align="center" width="100%" bgcolor="#FFFFFF">��</td>
</tr>
<form method="post" action="depadd.asp" name="name">
<tr> 
	<td align="center" colspan="4" width="100%"  height="34"> 
		���ӵ�λ��<input class=text type="text" name="depname" size="15" onMouseOver="window.status='����������Ҫ���ӵĵ�λ����';return true;" onMouseOut="window.status='';return true;" title="����������Ҫ���ӵĵ�λ����" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp;
		��λ���ţ�<input class=text type="text" name="deptype" size="15" onMouseOver="window.status='����������Ҫ���ӵĵ�λ����';return true;" onMouseOut="window.status='';return true;" title="����������Ҫ���ӵĲ�������" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp;

		<input type="submit" name="Submit" value="���" title="�������ť��������λ" onMouseOver="window.status='�������ť�������С��';return true;" onMouseOut="window.status='';return true;" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<input type="reset" value="��д" name="B1" title="�������ť������ӵ�λ" onMouseOver="window.status='�������ť������ӵ�λ';return true;" onMouseOut="window.status='';return true;" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>

</form>

</table>
</center>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
	<%conn.close
	set conn=nothing
end if%>