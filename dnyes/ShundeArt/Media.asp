<HTML>
<HEAD>
<TITLE>插入MediaPlay文件</TITLE>
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
        <legend>MEDIA定义</legend> 
        <table width="344" border="0" cellspacing="1" cellpadding="0"> 
          <tr> 
            <td width="54" height="25">文件类型</td> 
            <td height="25" colspan="3"><IFRAME name=ad src="upme5.htm" frameBorder=0 width="100%" scrolling=no 
height=20></IFRAME></td> 
          </tr> 
          <tr> 
            <td height="25">文件地址</td> 
            <td height="25" colspan="3"><input id=w name=w value='' style='width:200' title='可直接输入网上MediaPlay文件的地址，或由上传程序自动产生MediaPlay文件地址' size="20"> 
              avi、wmv、asf</td> 
          </tr> 
          <tr> 
            <td height="25">文件注释</td> 
            <td height="25" colspan="3"><input id=h name=h style='width:200' title='该注释将自动显示在文件的下方' size="20"></td> 
          </tr> 
          <tr> 
            <td height="25">立即播放</td> 
            <td width="84" height="25"><select name="select" id=b> 
                <option value="true">是
                <option value="false">否
              </select></td> 
            <td width="49" height="25">循环播放</td> 
            <td width="142"><select name="select" id=c> 
                <option value="-1">是
                <option value="0">否
              </select></td> 
          </tr> 
          <tr> 
            <td height="25">显示菜单</td> 
            <td height="25"><select name="select" id=d> 
                <option value="true">是
                <option value="false">否
              </select></td> 
            <td height="25">文件位置</td> 
            <td height="25"><select name="select" id=g> 
                <option value="left">居左
                <option value="center">居中
                <option value="right">居右
              </select></td> 
          </tr> 
          <tr> 
            <td height="25">播放高度</td> 
            <td height="25"><input name="Input" id=f style='width:50' ONKEYPRESS="event.returnValue=IsDigit();" value='240' size="20"></td> 
            <td height="25">播放宽度</td> 
            <td height="25"><input name="Input" id=e style='width:50' ONKEYPRESS="event.returnValue=IsDigit();" value='350' size="20"></td> 
          </tr> 
        </table> 
        </fieldset></td> 
      <td>&nbsp;</td> 
      <td><input type=button value=' 插 入 ' id=Ok title='点击“插入”按钮，在编辑器中插入该MediaPlay文件'> 
        <br> 
        <br> 
        <input class=button type=button onClick='window.close();' value=" 取 消 " name="button"></td> 
    </tr> 
  </table> 
</form> 
