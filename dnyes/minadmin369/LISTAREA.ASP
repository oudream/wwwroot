<!--#include file="SHEI45FECSA.ASP" -->
<%on error resume next%>
<%
sarea=request("sarea")
bigarea=request("bigarea")
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
    <td height="27">.:: 您可以在这里进行商品类别关操作</td>
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
      <div align="left">您现在操作的一级分类为 <b><%=sarea%></b>,二级分类为 <b><%=bigarea%></b>&nbsp;&nbsp;&nbsp; 
        [ <a href="addarea.asp?sarea=<%=sarea%>&bigarea=<%=bigarea%>&a=a">添加一个三级分类</a>]&nbsp;&nbsp;&nbsp;</div>
    </td>
  </tr>
  <tr>
    <td height="22" bgcolor="#FFCC33">三级分类名称</td>
    <td height="22" bgcolor="#FFCC33">编号</td>
    <td colspan="3" bgcolor="#FFCC33">操作</td>
  </tr>
  <%
sql="select * from area where sarea='"&sarea&"' and bigarea='"&bigarea&"' order by id desc"
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
for i=1 to rs.recordcount
%> 
  <tr>
    <td height="22" bgcolor="#FFFF99" width="45%"><%=rs("area")%></td>
    <td height="22" bgcolor="#FFFF99" width="10%"><%=rs("anumber")%></td>
    <td width="15%" bgcolor="#FFFF99">
<a href="addarea.asp?delarea=<%=rs("area")%>&sarea=<%=rs("sarea")%>&bigarea=<%=bigarea%>&a=d">删除</a></td>
    <td width="15%" bgcolor="#FFFF99">
<a href="addarea.asp?modiarea=<%=rs("area")%>&sarea=<%=rs("sarea")%>&bigarea=<%=bigarea%>&a=m">修改</a></td>
    <td width="15%" bgcolor="#FFFF99">
<a href="addarea.asp?unitarea=<%=rs("area")%>&sarea=<%=rs("sarea")%>&bigarea=<%=bigarea%>&a=u">合并</a></td>
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
