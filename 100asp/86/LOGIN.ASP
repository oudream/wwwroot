<html>
<head>
<title>网上购物</title>
</head>
<body bgcolor=lightgreen>
<div align="center">
<p align=center><FONT face=华文行楷 size=6> 
  实例90&nbsp;&nbsp; 用户登录</FONT>   </p>
<script language="VBScript">
<!--
sub submit1_onclick
if frm1.txt1.value="" then
msgbox "姓名不得为空,请重新输入！",0+16,"Warning"
exit sub
end if 
if frm1.txt2.value="" then
msgbox "密码不能为空，请重新输入！",0+16,"Warning"
exit sub
end if 
frm1.submit
end sub
-->
</script>
<form method="post" action="checklogin.asp" name="frm1">
      <TABLE border=1 cellPadding=1 cellSpacing=1 width="300">
           <tr><td colspan=2 align=middle><FONT face=华文新魏 size=+2>个人信息表</FONT></td></tr>
        <TR>
          <TD>姓名：</TD><TD><INPUT name=txt1></TD></tr>
        <TR>
          <TD>password：</TD><TD><INPUT type="password" name=txt2></TD>
        <tr><td colspan=2 align=middle><INPUT name=submit1 type=button value=用户登录>
    </td></tr></TABLE>
</div></FORM>
</body>
</html>
