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
    <td height="27">.:: 您可以在这里进行代理级别相关操作</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>
<table border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" width="60%" align="center">
  <tr> 
    <td height="22" bgcolor="#FFCC33" colspan="5"> 
      <div align="left"><a href="addlevel.asp?a=a">添加一个代理级别</a></div>
    </td>
  </tr>
  <tr> 
    <td height="22" bgcolor="#FFCC33">代理级别名称</td>
    <td height="22" bgcolor="#FFCC33">享受折扣</td>
    <td colspan="3" bgcolor="#FFCC33">操作</td>
  </tr>
<%
sql="select * from userlevel order by id desc"
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
for i=1 to rs.recordcount
%> 
  <tr> 
    <td height="22" bgcolor="#FFFF99" width="55%"><%=rs("userlevel")%></td>
    <td height="22" bgcolor="#FFFF99" width="15%"><%=rs("discount")%> %</td>
    <td width="15%" bgcolor="#FFFF99"><a href="addlevel.asp?dellevel=<%=rs("userlevel")%>&a=d">删除</a></td>
    <td width="15%" bgcolor="#FFFF99"><a href="addlevel.asp?modilevel=<%=rs("userlevel")%>&discount=<%=rs("discount")%>&a=m">修改</a></td>
  </tr>
  <%
rs.movenext
next
rs.close
set rs=nothing
%> 
</table>
<!--#include file="COPYRIGHT.ASP" -->
</body>
</html>
