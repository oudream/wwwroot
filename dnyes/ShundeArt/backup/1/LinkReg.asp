<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<%if linkreg="1" then%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
	<%if request.cookies(Forcast_SN)("key")="super" then%>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - �����������</title>
	<%else%>
<title><%=jjgn%> - ������������</title>
	<%end if%>
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
<!--#include file=User_Top.asp-->
<body topmargin="0">
<div align="center">
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="760" id="AutoNumber1">
<form align="center" method="post" name="FrmAddLink" LANGUAGE="javascript" onsubmit="return FrmAddLink_onsubmit()" action="<%if request.cookies(Forcast_SN)("key")="super" then%>SaveUserReg.asp<%else%>SaveReg.asp<%end if%>">
<tr class="TDtop">
	<td colspan=2 width="100%" height="25" align="center">�� <strong>�� �� �� �� �� ��</strong> ��</td>
</tr>
<tr>
	<td width="45%" height="32" align="right"> 
			<span class="smallFont">��&nbsp;&nbsp;&nbsp; �ͣ�</span>
	</td>
	<td width="55%" height="32"> 
		<select size="1" name="linktype" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" >
			<option selected value="">��ѡ����������</option>
			<option value="2">��ҳLOGO����</option>
			<option value="1">����LOGO����</option>
			<option value="0">��������</option>
		</select>
	</td>
</tr>
<tr>
	<td height="32" align="right" valign="middle"> 
		<span class="smallFont">��վ���ƣ�</span>
	</td>
	<td height="32"> 
		<div align="left"> 
			<input name="webname" size="26" class="smallInput" maxlength="20" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"  title="����������������վ���ƣ����Ϊ20������">
		</div>
	</td>
</tr>
<tr>
	<td height="32" align="right"> 
		<span class="smallFont">��վ��ַ��</span>
	</td>
	<td height="32"> 
		<div align="left"> 
			<input name="weburl" size="26" class="smallInput" maxlength="100" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" type="text"  value="http://" title="����������������վ��ַ�����Ϊ50���ַ���ǰ������http://">
		</div>
	</td>
</tr>
<tr>
	<td height="32" align="right"> 
		<span class="smallFont">��վLOGO��</span>
	</td>
	<td height="32"> 
		<div align="left"> 
			<input name="logo" size="26" class="smallInput" maxlength="100" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" type="text"  value="http://" title="����������������վLOGO��ַ�����Ϊ50���ַ���������ڵ�һѡ��ѡ������������ӣ�����Ͳ�����">
		</div>
	</td>
</tr>
<tr>
	<td height="32" align="right"> 
		<span class="smallFont">վ&nbsp;&nbsp;&nbsp;����</span>
	</td>
	<td height="32"> 
		<div align="left"> 
			<input name="webmaster" size="26" class="smallInput" maxlength="20" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" type="text"  title="�������������Ĵ����ˣ���Ȼ��֪������˭�������Ϊ20���ַ�">
		</div>
	</td>
</tr>
<tr> 
	<td height="32" align="right"> 
		<span class="smallFont">�����ʼ���</span>
	</td>
	<td height="32"> 
		<div align="left"> 
			<input name="email" size="26" class="smallInput" maxlength="30" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" type="text"  value="@" title="����������������ϵ�����ʼ������Ϊ30���ַ�">
		</div>
	</td>
</tr>
<tr>
	<td height="32" align="right"> 
		<span class="smallFont">��վ��飺</span>
	</td>
	<td height="32" valign="middle"> 
		<textarea rows="5" name="content" cols="29" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"  title="����������������վ�ļ򵥽���"></textarea>
	</td>
</tr>
<tr>
	<td colspan=2>
	<div align="center">
		<input type="hidden" name="pass" value="<%if request.cookies(Forcast_SN)("key")="super" then%>1<%else%>0<%end if%>">
		<input type="submit" value=" ȷ �� " name="cmdOk" class="buttonface" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp; 
		<input type="reset" value=" �� �� " name="cmdReset" class="buttonface" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</div>
	</td>
</tr>
</form>
</table>
</div>
</body>
</html>
<!--#include file=User_Bottom.asp-->
<%else
		Show_Err("�Բ��������������빦��δ������Ա����!<br><br><a href='javascript:history.back()'>����</a>")
		response.end
end if%>