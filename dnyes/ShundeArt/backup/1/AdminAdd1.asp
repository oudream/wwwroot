<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="char.inc"-->
<!--#include file="config.asp"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp" -->
<!--#include file="ChkManage.asp" -->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" and request.cookies(Forcast_SN)("ManagePurview")="99999") then
	Show_Err("�Բ������ĺ�̨����Ȩ�޲���������")
	response.end
else
	%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - ��ӹ����û�</title>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<link rel="stylesheet" type="text/css" href="site.css">
<script LANGUAGE="javascript">
<!--
function FrmAddLink_onsubmit() {
if (document.FrmAddLink.username.value=="")
	{
	  alert("�Բ�������������û�����")
	  document.FrmAddLink.username.focus()
	  return false
	 }
else if (document.FrmAddLink.username.value.length < 2)
	{
	  alert("�����û����ܲ��ܳ�һ�㣡")
	  document.FrmAddLink.username.focus()
	  return false
	 }
else if (document.FrmAddLink.username.value.length > 30)
	{
	  alert("�����û���̫���˰ɣ�")
	  document.FrmAddLink.username.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd.value=="")
	{
	  alert("�Բ������������룡")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd.value.length < 4)
	{
	  alert("Ϊ�˰�ȫ������Ӧ�ó�һ�㣡")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd.value.length > 16)
	{
	  alert("����̫���˰ɣ�")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.username.value==document.FrmAddLink.passwd.value)
	{
	  alert("Ϊ�˰�ȫ�������û��������벻Ӧ����ͬ��")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd2.value=="")
	{
	  alert("�Բ�������������֤���룡")
	  document.FrmAddLink.passwd2.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd2.value !== document.FrmAddLink.passwd.value)
	{
	  alert("�Բ�����������������벻һ�£�")
	  document.FrmAddLink.passwd2.focus()
	  return false
	 }
else if (document.FrmAddLink.fullname.value=="")
	{
	  alert("�Բ���������ù����û�����ʵ������")
	  document.FrmAddLink.fullname.focus()
	  return false
	 }
else if (document.FrmAddLink.depid.value=="")
	{
	  alert("�Բ�����ѡ��ù����û��Ĺ�����λ��")
	  document.FrmAddLink.depid.focus()
	  return false
	 }
else if (document.FrmAddLink.select.value=="")
	{
	  alert("�Բ�����ѡ��ù����û���Ȩ�ޣ�")
	  document.FrmAddLink.select.focus()
	  return false
	 }
}

//-->
</script>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<div align="center">
<table align=center border="1" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="<%=border%>">
<form align="center" method="post" name="FrmAddLink" LANGUAGE="javascript" onsubmit="return FrmAddLink_onsubmit()" action="AdminAdd2.asp">
<tr>
	<td width="100%" height=32 colspan="2" class="TDtop"><p align="center" >��<strong> �� �� �� �� �� �� �� </strong>��</td>
</tr>
<tr> 
	<td width="46%" height="32" align="right" valign="middle"> 
		<div align="center"><span class="smallFont">�� �� �� �� ����</span></div>
	</td>
	<td width="54%" height="32"> 
		<div align="left"> <input name="username" size="26" class="smallInput" maxlength="30" style="font-family: ����; font-size: 9pt" ></div>
	</td>
</tr>
<tr>
	<td width="46%" height="32" align="right"> 
		<div align="center"><span class="smallFont">��&nbsp;&nbsp;&nbsp; �룺</span></div>
	</td>
	<td width="54%" height="32"> 
		<div align="left"><input name="passwd" size="26" class="smallInput" maxlength="15" style="font-family: ����; font-size: 9pt" type="password" ></div>
	</td>
</tr>
<tr>
	<td width="46%" height="32" align="right"> 
		<div align="center"><span class="smallFont">��֤���룺</span></div>
	</td>
	<td width="54%" height="32"> 
		<div align="left"><input name="passwd2" size="26" class="smallInput" maxlength="15" style="font-family: ����; font-size: 9pt" type="password" ></div>
	</td>
</tr>
<tr> 
	<td width="46%" height="32" align="right"> 
		<div align="center"><span class="smallFont">��ʵ������</span></div>
	</td>
	<td width="54%" height="32"> 
		<div align="left"><input name="fullname" size="26" class="smallInput" maxlength="10" style="font-family: ����; font-size: 9pt" ></div>
	</td>
</tr>
<tr> 
	<td width="46%" height="32" align="right">
		<div align="center"><span class="smallFont">��Ӧǰ̨��ͨ�����û���</span></div>
	</td>
	<td width="54%" height="32"> 
		<div align="left"><input name="adder" size="26" class="smallInput" maxlength="10" style="font-family: ����; font-size: 9pt" >�����ʹ�������û��޷���½��</div>
	</td>
</tr>
<tr>
	<td width="46%" height="32" align="right"> 
		<div align="center"><span class="smallFont">������λ��</span></div>
	</td>
	<td width="54%" height="32"> <% 
	set rstype=createobject("adodb.recordset")
	sql="select * from "& db_Dep_Table &" order by id"
	rstype.Open sql,conn,1,3
	%>
		<select name="depid" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
			<option value="">��ѡ������λ</option>
	<%do while not rstype.EOF%>
			<option value="<%=rstype("id")%>"><%=rstype("depname")%>==<%=rstype("deptype")%></option>
		<%
		rstype.MoveNext
	loop
	set rstype=nothing
	%>
		</select>
	</td>
</tr>
<tr> 
	<td width="46%" height="32" align="right"> 
		<div align="center">�����û�Ȩ�ޣ�</div>
	</td>
	<td width="54%" height="32">
		<select name="purview" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
			<option  selected value="">��ѡ���û�Ȩ��</option>
			<option value=99999 selected>��������Ա</option>
			<option value=1>ϵͳ����Ա</option>
		</select>
	</td>
</tr>
<tr> 
	<td width="46%" colspan="2" height="32" align="right">&nbsp;</td>
</tr>
<tr>
	<td width="100%" colspan="2" height=32><p align="center">
		<input name="shenhe" type="hidden" value="0">
		<input name="oskey" type="hidden" value="super">
		<input type="button" value=" �� �� " onclick="javascript:history.go(-1)" class="unnamed5"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp;           
		<input type="submit" value=" ȷ �� " name="cmdOk" class="buttonface" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp; 
		<input type="reset" value=" �� �� " name="cmdReset" class="buttonface" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
</form>
</table>
</div>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
<%END IF%>
