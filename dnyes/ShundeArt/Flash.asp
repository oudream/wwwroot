<HTML><HEAD><TITLE>����Flash����</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<link rel="stylesheet" type="text/css" href="site.css">
<SCRIPT event=onclick for=Ok language=JavaScript>
	window.returnValue = a.value+"*"+b.value+"*"+c.value+"*"+d.value+"*"+e.value+"*"+f.value+"*"+g.value+"*"+h.value+"*"+j.value
	window.close();
</SCRIPT>
<script>
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
</script>
</HEAD>
<BODY bgColor=menu> 
<form name=flash> 
<br> 
<table border="0" cellspacing="1" cellpadding="0"> 
<tr> 
	<td>&nbsp;</td> 
	<td>
		<fieldset> 
		<legend>FLASH��������</legend> 
		<table width="329" border="0" cellpadding="0" cellspacing="1"> 
			<tr> 
				<td width="65" height="25" align="center" valign="middle">��������</td> 
				<td height="25" colspan="2"><iframe name="ad" frameborder=0 width=100% height=22 scrolling=no src=upme1.htm></iframe></td> 
			</tr> 
			<tr> 
				<td height="25" align="center" valign="middle">������ַ</td> 
				<td height="25" colspan="2">
					<input id=a name=a value=''  title='��ֱ����������Flash�����ĵ�ַ�������ϴ������Զ�����Flash������ַ' size="20">
				</td> 
			</tr> 
			<tr> 
				<td height="25" align="center" valign="middle">����ע��</td> 
				<td height="25" colspan="2">
					<input id=h name=h  title='��ע�ͽ��Զ���ʾ�ڶ������·�' size="20"> 
					<input id=j type="hidden" name=j value=''  title='��ֱ����������Flash�����ĵ�ַ�������ϴ������Զ�����Flash������ַ' size="20">
				</td> 
			</tr> 
			<tr> 
				<td height="25" align="center" valign="middle">��������</td> 
				<td height="25" colspan="1">
					<select name="select" id=b> 
						<option value="-1">��
						<option value="0">��
					</select>
					</td>
					<td height="25" colspan="1">
					ѭ������
					<select id=c>
						<option value="-1">��
						<option value="0">��
					</select>
					<br>
				</td> 
			</tr> 
			<tr> 
				<td height="25" align="center" valign="middle">��ʾ�˵�</td> 
				<td height="25">
					<select name="select" id=d > 
						<option value="-1">��
						<option value="0">��
					</select>
				</td> 
				<td height="25">�ļ�λ��
					<select name="select" id=g > 
						<option value="left">����
						<option value="center">����
						<option value="right">����
					</select>
				</td> 
			</tr> 
			<tr> 
				<td height="25" align="center" valign="middle">�����߶�</td> 
				<td width="62" height="25">
					<input name="Input" id=f  ONKEYPRESS="event.returnValue=IsDigit();" value='375' size="10">
				</td> 
				<td width="198" height="25">�������
					<input name="Input" id=e  ONKEYPRESS="event.returnValue=IsDigit();" value='500'  size="10">
				</td> 
			</tr> 
			</table> 
			</fieldset>
		</td> 
		<td>&nbsp;</td> 
		<td align="center" valign="middle">
			<input type=button value=' �� ��  ' id=Ok title='��������롱��ť���ڱ༭���в����Flash����' > 
			<br> 
			<br> 
			<br> 
			<input class=button type=button onClick='window.close();' value=" ȡ���� " name="button">
		</td> 
	</tr> 
</table> 
</form> 
</html>