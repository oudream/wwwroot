<!--#include file="SHEI45FECSA.ASP" -->
<%'on error resume next%>
<%
'开始添加操作
if request("action")="add" then
addlevel=request("addlevel")
discount=request("discount")

'检验此代理名称是否唯一
sql="select * from userlevel where userlevel='"&addlevel&"' order by id desc"
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then
response.write"<SCRIPT language=JavaScript>alert('对不起，此代理名称不唯一,请核实！');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end
end if
rs.close
set rs=nothing
'添加代理级别
conn.execute "insert into userlevel (userlevel,discount) values ('"&addlevel&"','"&discount&"')"
'conn.execute "insert into userlevel (level) values ('"&addlevel&"')"
response.redirect "discount.asp"
response.end
end if
%>
<%
'开始删除操作
if request("action")="del" then
conn.execute "delete * from userlevel where userlevel='"&request("dellevel")&"'"
response.redirect "discount.asp"
response.end
end if
%>
<%
'开始修改操作
if request("action")="modi" then

modilevel=request("modilevel")
modilevelto=request("modilevelto")
discount=request("discount")

'检验修改后的代理名称是否唯一
if modilevel<>modilevelto then
sql="select * from userlevel where userlevel='"&modilevelto&"' order by id desc"
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then
response.write"<SCRIPT language=JavaScript>alert('对不起，此代理名称不唯一,请核实！');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end
end if
rs.close
set rs=nothing
end if

conn.execute "update userlevel set userlevel='"&modilevelto&"',discount='"&discount&"' where userlevel='"&modilevel&"'"

response.redirect "discount.asp"
response.end
end if
%>
<html>
<head>
<title>商城系统管理中心</title>
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
<br>
<%if request("a")="a" then%>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" align="center" width="30%">
  <tr bgcolor="#FFCC33"> 
    <td colspan="2">添加代理级别</td>
  </tr>
  <form name="addlevel" method="post" action="addlevel.asp">
    <tr bgcolor="#FFFF99"> 
      <td align="right" width="30%">级别名称:</td>
      <td> 
        <input type="text" name="addlevel" class="form">
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td align="right">享受折扣:</td>
      <td> 
        <input type="text" name="discount" class="form">
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td colspan="2" align="center"> 
        <input type="submit" name="Submit" value="添 加">
        <input type="hidden" name="action" value="add">
      </td>
    </tr>
  </form>
</table>
<%end if%>
<br>
<%if request("a")="d" then%>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" align="center" width="30%">
  <tr bgcolor="#FFCC33"> 
    <td colspan="2">删除代理级别</td>
  </tr>
  <form name="dellevel" method="post" action="addlevel.asp">
    <tr bgcolor="#FFFF99"> 
      <td align="right" width="30%">级别名称:</td>
      <td> 
        <input type="text" name="dellevel" class="form" value=<%=request("dellevel")%> readonly>
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td colspan="2" align="center"> 
        <input type="submit" name="Submit2" value="删 除">
        <input type="hidden" name="action" value="del">
      </td>
    </tr>
  </form>
</table>
<%end if%>
<br>
<%
if request("a")="m" then
%>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" align="center" width="30%">
  <tr bgcolor="#FFCC33"> 
    <td colspan="2">修改代理级别</td>
  </tr>
  <form name="modilevel" method="post" action="addlevel.asp">
    <tr bgcolor="#FFFF99"> 
      <td align="right" width="30%">级别名称:</td>
      <td> 
        <input type="text" name="modilevelto" class="form" value=<%=request("modilevel")%>>
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td align="right">享受折扣:</td>
      <td> 
        <input type="text" name="discount" class="form" value=<%=request("discount")%>>
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td colspan="2" align="center"> 
        <input type="submit" name="Submit22" value="修 改">
        <input type="hidden" name="modilevel" value=<%=request("modilevel")%>>
        <input type="hidden" name="action" value="modi">
      </td>
    </tr>
  </form>
</table>
<%end if%>
<!--#include file="COPYRIGHT.ASP" -->
</body>
</html>
