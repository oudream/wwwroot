<!--#include file="SHEI45FECSA.ASP" -->
<%on error resume next%>
<%
'开始添加操作
if request("action")="add" then
addarea=request("addarea")
number=request("number")
photo=request("photo")

introt=replace(request("introt"),">","&gt;")
introt=replace(introt,"<","&lt;")
introt=replace(introt,"'","’")
if introt="" then introt="无"

introf=replace(request("introf"),">","&gt;")
introf=replace(introf,"<","&lt;")
introf=replace(introf,"'","’")
if introf="" then introf="无"

'检验此分类是否唯一
sql="select * from area where area='"&addarea&"' order by id desc"
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then
response.write"<SCRIPT language=JavaScript>alert('对不起，此分类名[或编号]不唯一,请核实！');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end
end if
rs.close
set rs=nothing
'添加类别
conn.execute "insert into area (area,anumber,photo,introt,introf) values ('"&addarea&"','"&number&"','"&photo&"','"&introt&"','"&introf&"')"
'如果类别被推荐到首页
if request("tuijian")="true" then
sql="select * from area order by id desc"
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
id=rs("id")
rs.close
set rs=nothing
conn.execute "update area set tuijian=true where id="&id
end if
response.redirect "area.asp"
response.end
end if
%>
<%
'开始删除操作
if request("action")="del" then
conn.execute "delete * from area where area='"&request("delarea")&"'"
conn.execute "delete * from subs where area='"&request("delarea")&"'"
response.redirect "area.asp"
response.end
end if
%>
<%
'开始修改操作
if request("action")="modi" then

modiarea=request("modiarea")
modiareato=request("modiareato")
photo=request("photo")

introt=replace(request("introt"),">","&gt;")
introt=replace(introt,"<","&lt;")
introt=replace(introt,"'","’")
if introt="" then introt="无"

introf=replace(request("introf"),">","&gt;")
introf=replace(introf,"<","&lt;")
introf=replace(introf,"'","’")
if introf="" then introf="无"
'检验此分类是否唯一
if modiarea<>modiareato then
sql="select * from area where area='"&modiareato&"' order by id desc"
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then
response.write"<SCRIPT language=JavaScript>alert('对不起，此分类名[或编号]不唯一,请核实！');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end
end if
rs.close
set rs=nothing
end if

conn.execute "update area set area='"&modiareato&"',photo='"&photo&"',introt='"&introt&"',introf='"&introf&"' where area='"&modiarea&"'"
conn.execute "update subs set area='"&modiareato&"' where area='"&modiarea&"'"


if request("tuijian")="true" then
conn.execute "update area set tuijian=true where area='"&modiareato&"'"
else
conn.execute "update area set tuijian=false where area='"&modiareato&"'"
end if
response.redirect "area.asp"
response.end
end if
%>
<%
'开始合并操作
if request("action")="unit" then

unitarea1=request("unitarea1")
unitarea2=request("unitarea2")
unitareato=request("unitareato")

if unitarea1="" or unitarea2="" or unitarea1=unitarea2 or unitareato="" then
response.write"<SCRIPT language=JavaScript>alert('对不起，您在合并分类时有错误操作,请核实！');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end
end if
conn.execute "update area set area='"&unitareato&"' where area='"&unitarea1&"'"
conn.execute "delete * from area where area='"&unitarea2&"'"
conn.execute "update subs set area='"&unitareato&"' where area='"&unitarea1&"' or area='"&unitarea2&"'"
response.redirect "area.asp"
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
    <td colspan="2">添加分类</td>
  </tr>
  <form name="addarea" method="post" action="addarea.asp">
    <%
sql="select * from area order by id desc"
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
if rs.eof or rs.bof then
number="0001"
else
number=rs("anumber")+1
number=right("0000"&number,4)
end if
rs.close
set rs=nothing
%> 
    <tr bgcolor="#FFFF99"> 
      <td align="right" width="25%">分类名:</td>
      <td> 
        <input type="text" name="addarea" class="form">
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td align="right">编号:</td>
      <td> 
        <input type="text" name="number" class="form" value=<%=number%> readonly>
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td align="right">图片:</td>
      <td> 
        <input type="text" name="photo" class="form">
      </td>
    </tr>
    <tr bgcolor="#FFFF99">
      <td align="right" valign="top">
        首页简介:</td>
      <td> 
        <textarea name="introt" class="form" cols="45" rows="5"></textarea></td>
    </tr>
    <tr bgcolor="#FFFF99">
      <td align="right" valign="top">
        分页简介:</td>
      <td> 
        <textarea name="introf" class="form" cols="45" rows="5"></textarea></td>
    </tr>
   <tr bgcolor="#FFFF99"> 
      <td align="right">状态:</td>
      <td> 是否在首页显示本分类 
        <input type="checkbox" name="tuijian" value="true">
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
    <td colspan="2">删除分类</td>
  </tr>
  <form name="delarea" method="post" action="addarea.asp">
    <tr bgcolor="#FFFF99"> 
      <td align="right" width="25%">分类名:</td>
      <td> 
        <input type="text" name="delarea" class="form" value=<%=request("delarea")%> readonly>
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
sql="select * from area where area='"&request("modiarea")&"'"
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
area=rs("area")
photo=rs("photo")
anumber=rs("anumber")
introt=rs("introt")
introf=rs("introf")
istuijian=""
if rs("tuijian")=True then istuijian="checked"
rs.close
set rs=nothing
%>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" align="center" width="50%">
  <tr bgcolor="#FFCC33"> 
    <td colspan="2">修改分类</td>
  </tr>
  <form name="modiarea" method="post" action="addarea.asp">
    <tr bgcolor="#FFFF99"> 
      <td align="right" width="25%">分类名:</td>
      <td> 
        <input type="text" name="modiareato" class="form" value=<%=area%>>
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td align="right">图片:</td>
      <td> 
        <input type="text" name="photo" class="form" value=<%=photo%>>
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td align="right">编号:</td>
      <td> 
        <input type="text" name="anumber" class="form" value=<%=anumber%> readonly>
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td align="right" valign="top"> 首页简介:</td>
      <td> 
        <textarea name="introt" class="form" cols="45" rows="5"><%=introt%></textarea>
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td align="right" valign="top"> 分页简介:</td>
      <td> 
        <textarea name="introf" class="form" cols="45" rows="5"><%=introf%></textarea>
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td align="right">状态:</td>
      <td> 是否在首页显示本分类 
        <input type="checkbox" name="tuijian" value="true" <%=istuijian%>>
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td colspan="2" align="center"> 
        <input type="submit" name="Submit22" value="修 改">
        <input type="hidden" name="modiarea" value=<%=request("modiarea")%>>
        <input type="hidden" name="action" value="modi">
      </td>
    </tr>
  </form>
</table>
<%end if%>
<%if request("a")="u" then%>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" align="center" width="30%">
  <tr bgcolor="#FFCC33"> 
    <td colspan="2">合并分类</td>
  </tr>
  <form name="unitarea" method="post" action="addarea.asp">
<tr bgcolor="#FFFF99"> 
      <td align="right" width="25%">分类名:</td>
      <td> 
        <input type="text" name="unitarea1" class="form" value=<%=request("unitarea")%> readonly>
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td align="right">分类名:</td>
      <td> 
        <select name="unitarea2">
          <option value="" selected>请选择一级分类(2)</option>
          <%
sql="select * from area where area<>'"&request("unitarea")&"'"
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
      <td align="right">合并为:</td>
      <td> 
        <input type="text" name="unitareato" class="form">
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td colspan="2" align="center"> 
        <input type="submit" name="Submit222" value="合 并">
        <input type="hidden" name="action" value="unit">
      </td>
    </tr>
  </form>
</table>
<%end if%>
<!--#include file="COPYRIGHT.ASP" -->
</body>
</html>
