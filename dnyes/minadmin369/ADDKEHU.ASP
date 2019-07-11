<!--#include file="SHEI45FECSA.ASP" -->
<%on error resume next%>

<%id=request("id")
'开始添加操作
if request("action")="add" then
kehu=request("kehu")
url=request("url")



'添加类别
conn.execute "insert into kehu (kehu,url) values ('"&kehu&"','"&url&"')"
'如果类别被推荐到首页
if request("tuijian")="true" then
sql="select * from kehu order by id desc"
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
id=rs("id")
rs.close
set rs=nothing
conn.execute "update kehu set kehutj=true where id="&id
end if
response.redirect "akehu.asp"
response.end
end if
%>
<%
'开始删除操作
if request("action")="del" then
conn.execute "delete * from kehu where id=" &id
response.redirect "akehu.asp"
response.end
end if
%>
<%
'开始修改操作
if request("action")="modi" then

kehu=request("kehu")
url=request("url")


conn.execute "update kehu set kehu='"&kehu&"',url='"&url&"' where id=" &id

if request("tuijian")="true" then
conn.execute "update kehu set kehutj=true where id=" &id
else
conn.execute "update kehu set kehutj=false where id=" &id
end if
response.redirect "akehu.asp"
response.end
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
    <td height="27">.:: 您可以在这里进行商品类别相关操作</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>
<br>
<%if request("a")="a" then%>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" align="center" width="50%">
  <tr bgcolor="#FFCC33"> 
    <td colspan="2">添加客户</td>
  </tr>
  <form name="addkehu" method="post" action="addkehu.asp">
    <tr bgcolor="#FFFF99"> 
      <td align="right" width="25%">客户名:</td>
      <td> 
        <input type="text" name="kehu" size=25 class="form">
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td align="right">URL:</td>
      <td> 
        <input type="text" name="url" class="form" size=25 value=http://>
      </td>
    </tr>
    
   <tr bgcolor="#FFFF99"> 
      <td align="right">状态:</td>
      <td> 是否在首页显示本客户 
        <input type="checkbox" name="tuijian" value="true">
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td colspan="2" align="center"> 
        <input type="submit" name="Submit" value="添 加">
        <input type="hidden" name="action" value="add"><input type="hidden" name="id" value=<%=id%>>
      </td>
    </tr>

  </form>
</table>
<%end if%>
<br>
<%if request("a")="d" then%>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" align="center" width="30%">
  <tr bgcolor="#FFCC33"> 
    <td colspan="2">删除分类</td>
  </tr>
<%
sql="select * from kehu where id=" &id
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
kehu=rs("kehu")
url=rs("url")
id=rs("id")
rs.close
set rs=nothing
%>
  <form name="delarea" method="post" action="addkehu.asp">
    <tr bgcolor="#FFFF99"> 
      <td align="right" width="25%">分类名:</td>
      <td> 
        <input type="text" name="delkehu" class="form" value=<%=kehu%> readonly>
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td colspan="2" align="center"> 
        <input type="submit" name="Submit2" value="删 除">
        <input type="hidden" name="action" value="del"><input type="hidden" name="id" value=<%=id%>>
      </td>
    </tr>
  </form>
</table>
<%end if%>
<br>
<%

if request("a")="m" then
sql="select * from kehu where id=" &id
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
kehu=rs("kehu")
url=rs("url")
istuijian=""
if rs("kehutj")=True then istuijian="checked"
rs.close
set rs=nothing
%>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" align="center" width="50%">
  <tr bgcolor="#FFCC33"> 
    <td colspan="2">修改客户</td>
  </tr>
  <form name="modiarea" method="post" action="addkehu.asp">
    <tr bgcolor="#FFFF99"> 
      <td align="right" width="25%">客户名:</td>
      <td> 
        <input type="text" name="kehu" size=25 class="form" value=<%=kehu%>>
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td align="right">URL:</td>
      <td> 
        <input type="text" name="url" size=25 class="form" value=<%=url%>>
      </td>
    </tr>
    
    <tr bgcolor="#FFFF99"> 
      <td align="right">状态:</td>
      <td> 是否在首页显示本客户 
        <input type="checkbox" name="tuijian" value="true" <%=istuijian%>>
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td colspan="2" align="center"> 
        <input type="submit" name="Submit22" value="修 改">
        <input type="hidden" name="action" value="modi"><input type="hidden" name="id" value=<%=id%>>
      </td>
    </tr>
  </form>
</table>
<%end if%>

<!--#include file="COPYRIGHT.ASP" -->
</body>
</html>
