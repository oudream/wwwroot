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
		%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - ���������޸�</title>
<script LANGUAGE="javascript">
<!--
function FrmAddLink_onsubmit() {
 if (document.FrmAddLink.linktype.value=="")
	{
	  alert("�Բ�����ѡ����������������ӵ����ͣ�")
	  document.FrmAddLink.linktype.focus()
	  return false
	 }
else if (document.FrmAddLink.webname.value=="")
	{
	  alert("�Բ���������������վ���ƣ�")
	  document.FrmAddLink.webname.focus()
	  return false
	 }
else if (document.FrmAddLink.weburl.value=="")
	{
	  alert("�Բ���������������վ��ַ��")
	  document.FrmAddLink.weburl.focus()
	  return false
	 }
else if (document.FrmAddLink.weburl.value=="http://")
	{
	  alert("�Բ���������������վ��ַ��")
	  document.FrmAddLink.weburl.focus()
	  return false
	 }
else if (document.FrmAddLink.webmaster.value=="")
	{
	  alert("�Բ��𣬲�֪������λ��")
	  document.FrmAddLink.webmaster.focus()
	  return false
	 }
else if (document.FrmAddLink.content.value=="")
	{
	  alert("�Բ���������������վ�ļ򵥽��ܣ�")
	  document.FrmAddLink.content.focus()
	  return false
	 }
else if (document.FrmAddLink.email.value=="")
	{
	  alert("�Բ������������ĵ����ʼ���")
	  document.FrmAddLink.email.focus()
	  return false
	 }
else if (document.FrmAddLink.email.value=="@")
	{
	  alert("�Բ������������ĵ����ʼ���")
	  document.FrmAddLink.email.focus()
	  return false
	 }
}

//-->
</script>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
		<%
		dim sql
		dim rs
		sql="select * from "& db_link_Table &" where id="&request("id")
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		%> 
<div align="center">
<table border="1" width="100%" cellpadding="0" cellspacing="0" align="center" bordercolor=<%=border%> style="border-collapse: collapse">
<form align="center" method="post" name="FrmAddLink" LANGUAGE="javascript" onsubmit="return FrmAddLink_onsubmit()" action="LinkSaveEdit.asp?id=<%=request("id")%>">
<tr class="TDtop">
	<td colspan=2 width="100%" height="25" align="center">�� <strong>�� �� �� �� �� ��</strong> ��</td>
</tr>
<tr>
	<td colspan=2 width="100%" height="25"></td>
</tr>
<tr> 
<tr>
	<td width="45%" height="32" align="right"> 
		<span class="smallFont">��&nbsp;&nbsp;&nbsp; �ͣ�</span>
	</td>
	<td width="55%" height="32"> 
		<select size="1" name="linktype" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
			<option <%if rs("linktype")="2" then%>selected<%end if%> value="2">��ҳLOGO����</option>
			<option <%if rs("linktype")="1" then%>selected<%end if%> value="1">����LOGO����</option>
			<option  <%if rs("linktype")="0" then%>selected<%end if%> value="0">��������</option>
		</select>
	</td>
</tr>
<tr> 
	<td height="32" align="right" valign="middle"> 
		<span class="smallFont">��վ���ƣ�</span>
	</td>
	<td height="32"> 
		<div align="left">
			<input name="webname" size="26" class="smallInput" maxlength="20" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="����������������վ���ƣ����Ϊ20������" value="<%=rs("webname")%>">
		</div>
	</td>
</tr>
<tr>
	<td height="32" align="right"> 
		<span class="smallFont">��վ��ַ��</span>
	</td>
	<td height="32"> 
		<div align="left"> 
			<input name="weburl" size="26" class="smallInput" maxlength="100" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"type="text" value="<%=rs("weburl")%>" title="����������������վ��ַ�����Ϊ50���ַ���ǰ������http://">
		</div>
	</td>
</tr>
<tr>
	<td height="32" align="right"> 
		<span class="smallFont">��վLOGO��</span>
	</td>
	<td height="32"> 
		<div align="left"> 
			<input name="logo" size="26" class="smallInput" maxlength="100" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"type="text"  value="<%=rs("logo")%>" title="����������������վLOGO��ַ�����Ϊ50���ַ���������ڵ�һѡ��ѡ������������ӣ�����Ͳ�����">
		</div>
	</td>
</tr>
<tr>
	<td height="32" align="right"> 
		<span class="smallFont">վ&nbsp;&nbsp;&nbsp;����</span>
	</td>
	<td height="32"> 
		<div align="left"> 
			<input name="webmaster" size="26" class="smallInput" maxlength="20" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"type="text"  title="�������������Ĵ����ˣ���Ȼ��֪������˭�������Ϊ20���ַ�" value="<%=rs("webmaster")%>">
		</div>
	</td>
</tr>
<tr> 
	<td height="32" align="right"> 
		<span class="smallFont">�����ʼ���</span>
	</td>
	<td height="32"> 
		<div align="left"> 
			<input name="email" size="26" class="smallInput" maxlength="30" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"type="text"  value="<%=rs("email")%>" title="����������������ϵ�����ʼ������Ϊ30���ַ�">
		</div>
	</td>
</tr>
<tr>
	<td height="32" align="right"> 
		<span class="smallFont">��վ��飺</span>
	</td>
	<td height="32" valign="middle"> 
		<textarea rows="5" name="content" cols="29" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="����������������վ�ļ򵥽���"><%=rs("content")%></textarea>
	</td>
</tr>
<tr>
	<td colspan=2 width="100%" height="25"></td>
</tr>
<tr class="TDtop">
	<td colspan=2 height="32" align="center"> 
		<div align="center">
			<input type="hidden" name="pass" value="1">
			<input type="submit" value=" ȷ �� " name="cmdOk" class="buttonface" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp;
			<input type="reset" value=" �� �� " name="cmdReset" class="buttonface" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		</div>
	</td>
</tr>
</form>
</table>
</center>
</div>
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