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
    <td height="27">.:: 您可以在这里进行商品相关操作</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" align="center" width="40%">
  <tr bgcolor="#FFCC33"> 
    <td colspan="2">请选择您要批量浏览的商品的条件</td>
  </tr>
  <tr bgcolor="#FFFFCC"> 
    <td colspan="2">由于查询全部商品过慢，所以请选择商品所在分类 </td>
  </tr>
  <form name="searchsubs" method="post" action="viewsubs.asp">
    <tr bgcolor="#FFFF99"> 
      <td align="right" width="35%">商品分类:</td>
      <td> 
        <select name="area">
          <option value="" selected>请选择商品所在分类</option>
          <%
sql="select * from area"
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
while not rs.eof
%> 
          <option value="<%=rs("area")%>"><%=rs("area")%></option>
          <%
rs.movenext
wend
rs.Close()
%> 
        </select>
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
