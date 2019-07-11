<!--#include file="SHEI45FECSA.ASP" -->
<%id=request("id")%>
<html>
<head>
<title>系统管理中心</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../inc/southdns.css" type="text/css">
</head>
<body bgcolor="#FFFFCC" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" vspace="0" hspace="0">
  <tr bgcolor="#FFCC33"> 
    <td height="27">.:: 您可以在这里进行商品图片资料相关操作</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>
<br>
<br>
<br>
<br>
<br>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" align="center" width="450">
  <tr bgcolor="#FFCC33"> 
    <td bgcolor="#FFCC33" width="15%" valign="top"><br>
      商品大图</td>
    <td width="85%"><br>
      <img src=showimage.asp?id=<%=id%>&tu=tu border="0"><br>
    </td>
  </tr>
  <tr bgcolor="#FFFF99"> 
    <td valign="top"><br>
      商品小图</td>
    <td><br>
      <img src=showimage.asp?id=<%=id%>&tu=stu border="0"><br>
    </td>
  </tr>
</table>
<!--#include file="COPYRIGHT.ASP" -->
</body>
</html>
