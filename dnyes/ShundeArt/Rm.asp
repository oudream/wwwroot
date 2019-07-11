
<HTML><HEAD><TITLE>插入RealPlay文件</TITLE>
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
      <legend>real文件定义</legend>
      <table width="346" border="0" cellspacing="1" cellpadding="0">
      <tr>
        <td width="69" height="25" align="center" valign="middle">文件类型</td>
        <td width="250" height="25"><IFRAME name=ad src="upme4.htm" frameBorder=0 width="100%" scrolling=no 
height=22></IFRAME></td>
      </tr>
      <tr>
        <td height="25" align="center" valign="middle">文件地址</td>
        <td height="25"><input name=r id=r style='width:200' title='可直接输入网上RealPlay文件的地址，或由上传程序自动产生RealPlay文件的地址' value='' size="20">
      rm、ram</td>
      </tr>
      <tr>
        <td height="25" align="center" valign="middle">文件注释</td>
        <td height="25"><input name=g id=g style='width:200' title='该注释将自动显示在文件的下方' size="20"></td>
      </tr>
      <tr>
        <td height="25" align="center" valign="middle">立即播放</td>
        <td height="25"><select name="select" id=b>
            <option value="-1">是
            <option value="0">否
        </select></td>
      </tr>
      <tr>
        <td height="25" align="center" valign="middle">文件位置</td>
        <td height="25"><select name="select" id=d>
            <option value="left">居左
            <option value="center">居中
            <option value="right">居右
        </select></td>
      </tr>
      <tr>
        <td height="25" align="center" valign="middle">播放高度</td>
        <td height="25"><input name="Input" id=e style='width:50' ONKEYPRESS="event.returnValue=IsDigit();" value='240' size="20">
          播放宽度：
            <input name="Input" id=f style='width:50' ONKEYPRESS="event.returnValue=IsDigit();" value='352' size="20"></td>
      </tr>
      <tr>
        <td height="25" align="center" valign="middle">循环播放</td>
        <td height="25"><select name="select" id=c>
            <option value="-1">是
            <option value="0">否
        </select></td>
      </tr>
    </table> 
    </fieldset> </td>
    <td>&nbsp;</td>
    <td>
        <input name="button" type=button id=Ok title='点击“插入”按钮，在编辑器中插入该RealPlay文件' value=' 插  入 '>
        <br>
        <br><input class=button type=button onClick='window.close();' value=" 取　消 " name="button">
        <br>
    </td>
  </tr>
</table>

</form>
