<!--#include file="SHEI45FECSA.ASP" -->
<!--#include file="../mdf.asp" -->
<%on error resume next%>
<%
if session("y_eshop_aleave")="u" then
response.write "<br><br><br><center><font size=-1>Sorry,您的权限不足,无法完成此操作 ^&^ .</font></center>"
response.end
end if
%>
<%on error resume next%>
<%
if request("action")="add" then
auid=request("auid")
apwd=request("apwd")
apwd=md5(apwd)
aleave=request("aleave")

conn.execute "insert into adminuser (auid,apwd,aleave) values ('"&auid&"','"&apwd&"','"&aleave&"')"
response.write"<SCRIPT language=JavaScript>alert('管理员添加成功！');"
response.write"javascript:history.go(-1)</SCRIPT>"
end if
%>
<%
if request("action")="del" then
conn.execute "delete * from adminuser where auid='"&request("auid")&"'"
response.write"<SCRIPT language=JavaScript>alert('管理员删除成功！');"
response.write"javascript:history.go(-1)</SCRIPT>"
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
    <td height="27">.:: 您可以在这里进行[管理员]相关操作</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>

<table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" align="center" width="250">
  <tr bgcolor="#FFCC33"> 
    <td colspan="2">添加管理员：</td>
  </tr>
  <form name="addadmin" method="post" action="ADDADMIN.ASP">
    <tr bgcolor="#FFFF99"> 
      <td align="right" bgcolor="#FFFF99" height="22">帐号:</td>
      <td> 
        <input type="text" name="auid" class="form">
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td align="right" bgcolor="#FFFF99" height="22">密码:</td>
      <td> 
        <input type="password" name="apwd" class="form">
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td colspan="2" align="center"> 
        <input type="submit" name="Submit" value="确 定">
        <input type="hidden" name="action" value="add">
        <input type="hidden" name="aleave" value="u">
      </td>
    </tr>
  </form>
</table>
<br>
<br>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" align="center" width="250">
  <tr bgcolor="#FFCC33"> 
    <td colspan="2">删除管理员：</td>
  </tr>
  <form name="deladmin" method="post" action="ADDADMIN.ASP">
    <tr bgcolor="#FFFF99"> 
      <td align="right" bgcolor="#FFFF99" height="22">帐号:</td>
      <td> 
        <select name="auid">
          <option value="" selected>将删除的管理员</option>
<%
sql="select * from adminuser where aleave<>'s'"
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
while not rs.eof
%> 
          <option value="<%=rs("auid")%>"><%=rs("auid")%></option>
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
        <input type="submit" name="Submit2" value="确 定">
        <input type="hidden" name="action" value="del">
      </td>
    </tr>
  </form>
</table>
<!--#include file="COPYRIGHT.ASP" -->
</body>
</html>
