
<HTML><HEAD><TITLE>����RealPlay�ļ�</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<link rel="stylesheet" type="text/css" href="site.css">
<SCRIPT event=onclick for=Ok language=JavaScript>
  window.returnValue = r.value+"*"+b.value+"*"+c.value+"*"+d.value+"*"+e.value+"*"+f.value+"*"+g.value
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
<form name=rm><br>
<table width="200" border="0" cellspacing="1" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
    <td><fieldset> 
      <legend>real�ļ�����</legend>
      <table width="346" border="0" cellspacing="1" cellpadding="0">
      <tr>
        <td width="69" height="25" align="center" valign="middle">�ļ�����</td>
        <td width="250" height="25"><IFRAME name=ad src="upme4.htm" frameBorder=0 width="100%" scrolling=no 
height=22></IFRAME></td>
      </tr>
      <tr>
        <td height="25" align="center" valign="middle">�ļ���ַ</td>
        <td height="25"><input name=r id=r style='width:200' title='��ֱ����������RealPlay�ļ��ĵ�ַ�������ϴ������Զ�����RealPlay�ļ��ĵ�ַ' value='' size="20">
      rm��ram</td>
      </tr>
      <tr>
        <td height="25" align="center" valign="middle">�ļ�ע��</td>
        <td height="25"><input name=g id=g style='width:200' title='��ע�ͽ��Զ���ʾ���ļ����·�' size="20"></td>
      </tr>
      <tr>
        <td height="25" align="center" valign="middle">��������</td>
        <td height="25"><select name="select" id=b>
            <option value="-1">��
            <option value="0">��
        </select></td>
      </tr>
      <tr>
        <td height="25" align="center" valign="middle">�ļ�λ��</td>
        <td height="25"><select name="select" id=d>
            <option value="left">����
            <option value="center">����
            <option value="right">����
        </select></td>
      </tr>
      <tr>
        <td height="25" align="center" valign="middle">���Ÿ߶�</td>
        <td height="25"><input name="Input" id=e style='width:50' ONKEYPRESS="event.returnValue=IsDigit();" value='240' size="20">
          ���ſ�ȣ�
            <input name="Input" id=f style='width:50' ONKEYPRESS="event.returnValue=IsDigit();" value='352' size="20"></td>
      </tr>
      <tr>
        <td height="25" align="center" valign="middle">ѭ������</td>
        <td height="25"><select name="select" id=c>
            <option value="-1">��
            <option value="0">��
        </select></td>
      </tr>
    </table> 
    </fieldset> </td>
    <td>&nbsp;</td>
    <td>
        <input name="button" type=button id=Ok title='��������롱��ť���ڱ༭���в����RealPlay�ļ�' value=' ��  �� '>
        <br>
        <br><input class=button type=button onClick='window.close();' value=" ȡ���� " name="button">
        <br>
    </td>
  </tr>
</table>

</form>
