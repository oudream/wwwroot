<HTML>
<HEAD>
<TITLE>����MediaPlay�ļ�</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<link rel="stylesheet" type="text/css" href="site.css">
<SCRIPT event=onclick for=Ok language=JavaScript>
  window.returnValue = w.value+"*"+b.value+"*"+c.value+"*"+d.value+"*"+e.value+"*"+f.value+"*"+g.value+"*"+h.value
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
<form name=media> 
  <br> 
  <table border="0" cellspacing="1" cellpadding="0"> 
    <tr align="center" valign="middle"> 
      <td>&nbsp;</td> 
      <td><fieldset> 
        <legend>MEDIA����</legend> 
        <table width="344" border="0" cellspacing="1" cellpadding="0"> 
          <tr> 
            <td width="54" height="25">�ļ�����</td> 
            <td height="25" colspan="3"><IFRAME name=ad src="upme5.htm" frameBorder=0 width="100%" scrolling=no 
height=20></IFRAME></td> 
          </tr> 
          <tr> 
            <td height="25">�ļ���ַ</td> 
            <td height="25" colspan="3"><input id=w name=w value='' style='width:200' title='��ֱ����������MediaPlay�ļ��ĵ�ַ�������ϴ������Զ�����MediaPlay�ļ���ַ' size="20"> 
              avi��wmv��asf</td> 
          </tr> 
          <tr> 
            <td height="25">�ļ�ע��</td> 
            <td height="25" colspan="3"><input id=h name=h style='width:200' title='��ע�ͽ��Զ���ʾ���ļ����·�' size="20"></td> 
          </tr> 
          <tr> 
            <td height="25">��������</td> 
            <td width="84" height="25"><select name="select" id=b> 
                <option value="true">��
                <option value="false">��
              </select></td> 
            <td width="49" height="25">ѭ������</td> 
            <td width="142"><select name="select" id=c> 
                <option value="-1">��
                <option value="0">��
              </select></td> 
          </tr> 
          <tr> 
            <td height="25">��ʾ�˵�</td> 
            <td height="25"><select name="select" id=d> 
                <option value="true">��
                <option value="false">��
              </select></td> 
            <td height="25">�ļ�λ��</td> 
            <td height="25"><select name="select" id=g> 
                <option value="left">����
                <option value="center">����
                <option value="right">����
              </select></td> 
          </tr> 
          <tr> 
            <td height="25">���Ÿ߶�</td> 
            <td height="25"><input name="Input" id=f style='width:50' ONKEYPRESS="event.returnValue=IsDigit();" value='240' size="20"></td> 
            <td height="25">���ſ��</td> 
            <td height="25"><input name="Input" id=e style='width:50' ONKEYPRESS="event.returnValue=IsDigit();" value='350' size="20"></td> 
          </tr> 
        </table> 
        </fieldset></td> 
      <td>&nbsp;</td> 
      <td><input type=button value=' �� �� ' id=Ok title='��������롱��ť���ڱ༭���в����MediaPlay�ļ�'> 
        <br> 
        <br> 
        <input class=button type=button onClick='window.close();' value=" ȡ �� " name="button"></td> 
    </tr> 
  </table> 
</form> 
