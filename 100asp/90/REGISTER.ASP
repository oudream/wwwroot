<html>
<head>
<title>聊天室</title>
</head>
<body bgcolor=lightgreen>
<div align="center">
<p align=center><FONT face=华文行楷 size=6> 
  实例99&nbsp;&nbsp; 聊天室注册表</FONT>   </p>
<form method="post" action="adduser.asp" name="form1">
      <TABLE border=1 cellPadding=1 cellSpacing=1 width="600" style="HEIGHT: 206px; WIDTH: 346px">
           <tr><td colspan=4 align=middle><FONT face=华文新魏 size=+2>个人信息表</FONT></td></tr>
        <TR>
          <TD>网友昵称：</TD>
          <TD><INPUT maxLength=8 name=txt1></TD><tr>
          <TD>性别：</TD>
          <TD><SELECT name=sel1>
				<OPTION selected>男</OPTION>
				<OPTION>女</OPTION></SELECT></TD></tr>
        <TR>
          <TD>输入密码：</TD>
          <TD><INPUT type="password" name=txt2></TD></tr>
        <tr>
          <TD>确认密码：</TD>
          <TD><INPUT type="password" name=txt3></TD></tr>
        <TR>
          <TD>E-mail: </TD><TD><INPUT name=txt4></TD></TR>
        <tr><td colspan=2 align=middle><INPUT type="submit" value="用户注册" id=submit1 name=submit1>
     </td></tr></TABLE></FORM>
</div>
</body>
</html>
