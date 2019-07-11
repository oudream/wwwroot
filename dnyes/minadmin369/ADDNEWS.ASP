<!--#include file="SHEI45FECSA.ASP" -->
<%on error resume next%>
<%
'添加公告
if request("action")="add" then
color=request("color")
if trim(color)="" then color=" "
subject=request("subject")
newstype=request("newstype")
message=request("message")
conn.execute "insert into message (subject,message,idate,newstype,color) values ('"&subject&"','"&message&"','"&date()&"','"&newstype&"','"&color&"')"
response.redirect "viewnews.asp"
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
    <td height="27">.:: 您可以在这里进行公告相关操作</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>
<table border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" width="65%" align="center">
  <form action="ADDNEWS.ASP" method="post" name="addnews">
    <tr bgcolor="#FFCC33"> 
      <td height="22" width="19%">&nbsp;公告标题：</td>
      <td colspan="3"> &nbsp;&nbsp;
        <input maxlength=45 name="subject" size=40 class="form"> </td>
    </tr>
    <tr bgcolor="#FFCC33"> 
      <td height="22" width="19%">公告类型：</td>
      <td width="22%"> &nbsp;&nbsp;
        <select name="newstype">
          <option value="行业信息">行业信息</option>
          <option value="本站动态">本站动态</option>
          <option value="客服中心">客服中心</option>
        </select> </td>
      <td width="27%">color</td>
      <td width="32%"><input maxlength=6 name="color" size=10 class="form"></td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td height="22" valign="top"><br> &nbsp;公告内容：</td>
      <td colspan="3"><br>
        &nbsp; <textarea name="message" cols="50" class="form" rows="8">
</textarea>
        <br>
        <br> </td>
    </tr>
    <tr bgcolor="#FFFFCC"> 
      <td height="45" colspan="4"> <div align="center"> 
          <input type="submit" name="Submit" value="确 定">&nbsp;&nbsp;
          <input type="reset" name="Submit" value="取 消">&nbsp;&nbsp;
          <input type="button" name="Submit" value="后 退" onClick="history.go( -1 );return true;">
          <input type="hidden" name="action" value="add">
        </div></td>
    </tr>
  </form>
</table>
<!--#include file="COPYRIGHT.ASP" -->
</body>
</html>
