<!--#include file="SHEI45FECSA.ASP" -->
<%on error resume next%>
<html>
<head>
<title>系统管理中心</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../inc/southdns.css" type="text/css">
</head>
<body bgcolor="#FFFFCC" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" vspace="0" hspace="0">
  <tr bgcolor="#FFCC33"> 
    <td height="27">.:: 您可以在这里进行用户相关操作</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" align="center">
  <tr bgcolor="#FFCC33"> 
    <td colspan="2">模糊查询</td>
  </tr>
  <form name="searchuser" method="post" action="VIEWUSER.ASP">
    <tr bgcolor="#FFFF99"> 
      <td align="right">用户名:</td>
      <td> 
        <input type="text" name="uid" class="form">
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td colspan="2" align="center"> 
        <input type="submit" name="Submit" value="确 定">
      </td>
    </tr>
</form>
</table>
<!--#include file="COPYRIGHT.ASP" -->
</body>
</html>
