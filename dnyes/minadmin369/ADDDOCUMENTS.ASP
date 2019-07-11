<!--#include file="SHEI45FECSA.ASP" -->
<%'on error resume next%>
<%
'添加公告
if request("action")="add" then
color=request("color")
if trim(color)="" then color=" "
subject=request("subject")
addr=request("addr")
conn.execute "insert into document (subject,addr,indate,color) values ('"&subject&"','"&addr&"','"&date()&"','"&color&"')"
response.redirect "ADDDOCUMENTSOK.ASP"
end if 
%>
<html>
<head>
<title>系统管理中心</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../inc/southdns.css" type="text/css">
</head>
<body bgcolor="#FFFFCC" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" vspace="0" hspace="0">
  <tr bgcolor="#FFCC33"> 
    <td height="27">.:: 您可以在这里进行添加文档操作</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>
<table border="0" cellspacing="1" cellpadding="5" bgcolor="#000000" width="65%" align="center">
  <form action="ADDDOCUMENTS.ASP" method="post" name="addnews">
    <tr bgcolor="#FFCC33"> 
      <td height="22" width="19%">文档标题：</td>
      <td> &nbsp;&nbsp; <input maxlength=45 name="subject" size=40 class="form"> 
      </td>
    </tr>
    <tr bgcolor="#FFCC33"> 
      <td height="22" width="19%">文档地址：</td>
      <td> &nbsp;&nbsp; <input maxlength=45 name="addr" size=40 class="form" value="mydocument/"></td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td height="22" valign="top">color</td>
      <td>&nbsp;&nbsp; <input maxlength=6 name="color" size=10 class="form"></td>
    </tr>
    <tr bgcolor="#FFFFCC"> 
      <td height="45" colspan="2"> <div align="center"> 
          <input type="submit" name="Submit" value="确 定">
          &nbsp;&nbsp; 
          <input type="reset" name="Submit" value="取 消">
          &nbsp;&nbsp; 
          <input type="button" name="Submit" value="后 退" onClick="history.go( -1 );return true;">
          <input type="hidden" name="action" value="add">
        </div></td>
    </tr>
  </form>
</table>
<!--#include file="COPYRIGHT.ASP" -->
</body>
</html>
