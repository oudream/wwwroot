<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
<head>
<title>信网 数信科技 - 用户注册成功</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="refresh" content="3;url=index.asp">
<link rel="stylesheet" href="css.css" type="text/css">
</head>

<body>
<p>&nbsp;</p><br>
<table width="600" border="0" align="center" cellpadding="0" cellspacing="0" >
  <tr> 
    <td height="23" align="center"><font color="#FF0000">注册成功！</font></td>
  </tr>
  <tr> 
    <td height="27" align="center"><div align="left">1、邮件成功发到你的邮箱</div></td>
  </tr>
  <tr> 
    <td height="27" align="center"><div align="left"></div></td>
  </tr>
  <tr> 
    <td height="28"> 2、多谢您对我们的支持，如有需要请联系我们</td>
  </tr>
  <tr> 
    <td height="35" align="center">&nbsp;</td>
  </tr>
  <tr>
    <td height="28">3、请您紧记您的帐号：<font color="#FF0000"><%=session("y_it_uid")%></font>；&nbsp;密码：<font color="#FF0000"><%=request("pwd")%>；&nbsp;邮箱：<font color="#FF0000"><%=session("y_it_usermail")%></font>。
</td>
  </tr>
</table>
<p>&nbsp;</p><br><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><div align="center">页面在<font color="#FF0000"> 3</font> 秒钟后自动刷新</div></td>
  </tr>
</table>
</body>
</html>
