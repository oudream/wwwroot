<html>
<head>
<title> 聊天室</title>
</head>
<body bgcolor=lightgreen>
<div align="center">
<p align=center><FONT face=华文行楷 size=6> 
  实例99&nbsp;&nbsp; 聊天室</FONT>   </p>
<script language="VBScript">
<!--
sub button1_onclick
	location.href="register.asp"
end sub
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
      <TABLE border=1 cellPadding=1 cellSpacing=1 width="600" style="HEIGHT: 123px; WIDTH: 267px">
           <tr><td colspan=4 align=middle><FONT face=华文新魏 size=5>用户登录 
      </FONT></td></tr>
        <TR>
          <TD>昵&nbsp;&nbsp; 称：</TD><TD><INPUT name=txt1></TD></TR>
        <tr>
          <TD>password：</TD><TD><INPUT type="password" name=txt2></TD>
        <tr><td colspan=2 align=middle>
        <INPUT name=submit1 type=button value=用户登录>
        <INPUT type="button" value="用户注册" id=button1>
    </td></tr></TABLE></form>
</div>
<DIV align=center><BR><FONT 
face=华文新魏><STRONG>使用此程序可以用昵称：<U>lyf</U>和密码：<U>2000</U>登录</STRONG></FONT></DIV>
</body>
</html>
