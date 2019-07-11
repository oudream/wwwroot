<html>
<head>
<title>留言本</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<script language="VBScript">
sub check()
if frm1.user.value=empty then
msgbox "请输入您的留言昵称！"
focusto(0)
exit sub
end if 
if frm1.email.value=empty then
msgbox "请输入您的电子邮件地址！"
focusto(1)
exit sub
end if 
if instr(frm1.email.value, "@")=0 then
  msgbox "请正确填写您的电子邮件地址！"
  focusto(1)
exit sub
end if
if frm1.content.value=empty then
msgbox "请输入您的留言内容！"
focusto(2)
exit sub
end if 
frm1.submit
end sub
sub focusto(x)
document.frm1.elements(x).focus()
end sub 
</script>
<body bgcolor="#e2f3f2" text="#333333">
<P align=center>
 <font face="华文行楷" size="+3"><font face="行楷">实例　留言本</font></font> </P>
<P align=center><b><A href="display.asp"><font size="+1">查看留言</font></A></b>
<table width="72%" border="0" align="center">
  <tr bgcolor="#e2f3f2" align="middle" valign="center"> 
    <td height="385"> 
 <form method="post" action="savadata.asp" name="frm1">
   <table align=center bgcolor=#e2f3f2 border=1 cellpadding=0 cellspacing=0 width="100%">
<tbody> 
<tr> 
  <td>留言者:</td>
  <td><input maxlength=30 name="user"></td>
</tr>
<tr> 
  <td>e_mail:</td>
  <td colspan=3><input name="email" >
  </td>
</tr>
<tr> 
  <td>标题:</td>
  <td colspan=3><input maxlength=200 name=subject size=40></td>
</tr>
<tr> 
  <td height="37">留给谁？</td>
  <td colspan=3><input name="towho" >
  </td>
</tr>
<tr> 
  <td>内容:</td>
  <td colspan=3><TEXTAREA cols=51 name=content rows=8></TEXTAREA>
  </td>
</tr>
<tr align=middle> 
  <td colspan=4 align=center><input name=B12  type=button value=" 提 交 " onclick="check()">
    <input name=B2  type=reset value=" 重写 "></td>
</tr>
</tbody> 
   </table>
 </form></td></tr></table></P>
</body>  
</html>  
