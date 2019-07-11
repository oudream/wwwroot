<!--#include file="SHEI45FECSA.ASP" -->
<%on error resume next%>
<%
'开始添加操作
if request("action")="add" then
addgys=request("addgys")
photo=request("photo")
icp=request("icp")
hongdun=request("hongdun")

if icp="" then icp="无"
if hongdun="" then hongdun="无"
gysjj=replace(request("gysjj"),">","&gt;")
gysjj=replace(gysjj,"<","&lt;")
gysjj=replace(gysjj,"'","’")
if gysjj="" then gysjj="无"

gysxq=replace(request("gysxq"),">","&gt;")
gysxq=replace(gysxq,"<","&lt;")
gysxq=replace(gysxq,"'","’")
if gysxq="" then gysxq="无"

'检验此分类是否唯一
sql="select * from gys where gys='"&addgys&"' order by id desc"
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
conn.execute "insert into gys (gys,photo,icp,hongdun,gysjj,gysxq) values ('"&addgys&"','"&photo&"','"&icp&"','"&hongdun&"','"&gysjj&"','"&gysxq&"')"
'如果类别被推荐到首页
if request("ifgystj")="y" then
sql="select * from gys order by id desc"
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
id=rs("id")
rs.close
set rs=nothing
conn.execute "update gys set ifgystj=true where id="&id
end if
response.redirect "gys.asp"
response.end
end if
%>
<%
'开始删除操作
if request("action")="del" then
conn.execute "delete * from gys where gys='"&request("delgys")&"'"
conn.execute "delete * from subs where gys='"&request("delgys")&"'"
response.redirect "gys.asp"
response.end
end if
%>
<%
'开始修改操作
if request("action")="modi" then

modigysto=request("modigysto")
modigys=request("modigys")
photo=request("photo")
icp=request("icp")
hongdun=request("hongdun")
id=request("id")

if icp="" then icp="无"
if hongdun="" then hongdun="无"
gysjj=replace(request("gysjj"),">","&gt;")
gysjj=replace(gysjj,"<","&lt;")
gysjj=replace(gysjj,"'","’")
if gysjj="" then gysjj="无"

gysxq=replace(request("gysxq"),">","&gt;")
gysxq=replace(gysxq,"<","&lt;")
gysxq=replace(gysxq,"'","’")
if gysxq="" then gysxq="无"
'检验此分类是否唯一
if modigys<>modigysto then
sql="select * from gys where gys='"&modigysto&"' order by id desc"
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then
response.write"<SCRIPT language=JavaScript>alert('对不起，此分类名不唯一,请核实！');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end
end if
rs.close
set rs=nothing
end if

conn.execute "update gys set gys='"&modigysto&"',photo='"&photo&"',icp='"&icp&"',hongdun='"&hongdun&"',gysjj='"&gysjj&"',gysxq='"&gysxq&"' where id=" & id
conn.execute "update subs set gys='"&modigysto&"' where gys='"&request("modigys")&"'"

if request("ifgystj")="y" then
conn.execute "update gys set ifgystj=true where id="&id
else
conn.execute "update gys set ifgystj=false where id="&id
end if

response.redirect "gys.asp"
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
    <td height="10">&nbsp;</td>
  </tr>
</table>
<%if request("a")="a" then%>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" align="center" width="50%">
  <tr bgcolor="#FFCC33"> 
    <td colspan="2">添加供应商</td>
  </tr>
  <form name="addarea" method="post" action="addgys.asp">
    <%
sql="select * from gys order by id desc"
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
rs.close
set rs=nothing
%> 
    <tr bgcolor="#FFFF99"> 
      <td align="left" width="25%">供应商:</td>
      <td> 
        <input type="text" name="addgys" class="form">
      </td>
    </tr><tr bgcolor="#FFFF99"> 
      <td align="left">图片:</td>
      <td> 
        <input type="text" name="photo" class="form">
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td align="left">ICP证:</td>
      <td> 
        <input type="text" name="icp" class="form">
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td align="left">红盾:</td>
      <td> 
        <input type="text" name="hongdun" class="form">
      </td>
    </tr>
    <tr bgcolor="#FFFF99">
      <td align="left" valign="top">
        首页简介:</td>
      <td> 
        <textarea name="gysjj" class="form" cols="45" rows="5"></textarea></td>
    </tr>
    <tr bgcolor="#FFFF99">
      <td align="left" valign="top">
        分页简介:</td>
      <td> 
        <textarea name="gysxq" class="form" cols="45" rows="5"></textarea></td>
    </tr>
   <tr bgcolor="#FFFF99"> 
      <td align="left">状态:</td>
      <td> 是否在首页显示本分类 
        <input type="checkbox" name="ifgystj" value="y">
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
<%if request("a")="d" then%>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" align="center" width="30%">
  <tr bgcolor="#FFCC33"> 
    <td colspan="2">删除分类</td>
  </tr>
  <form name="delarea" method="post" action="addgys.asp">
    <tr bgcolor="#FFFF99"> 
      <td align="right" width="25%">分类名:</td>
      <td> 
        <input type="text" name="delgys" class="form" value=<%=request("delgys")%> readonly>
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
sql="select * from gys where gys='"&request("modigys")&"'"
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
gys=rs("gys")
photo=rs("photo")
icp=rs("icp")
hongdun=rs("hongdun")
gysjj=rs("gysjj")
gysxq=rs("gysxq")
id=rs("id")
isifgystj=""
if rs("ifgystj")=True then isifgystj="checked"
rs.close
set rs=nothing
%>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" align="center" width="50%">
  <tr bgcolor="#FFCC33"> 
    <td colspan="2">供应商资料修改</td>
  </tr>
  <form name="modi" method="post" action="addgys.asp">
    <tr bgcolor="#FFFF99"> 
      <td align="left" width="25%">供应商:</td>
      <td> 
        <input type="text" name="modigysto" class="form" value=<%=gys%>>
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td align="left">图片:</td>
      <td> 
        <input type="text" name="photo" class="form" value=<%=photo%>>
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td align="left">ICP证:</td>
      <td> 
        <input type="text" name="icp" class="form" value=<%=icp%>>
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td align="left">红盾:</td>
      <td> 
         <input type="text" name="hongdun" class="form" value=<%=hongdun%>>
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td align="left" valign="top"> 首页简介:</td>
      <td> 
        <textarea name="gysjj" class="form" cols="45" rows="5"><%=gysjj%></textarea>
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td align="left" valign="top"> 分页简介:</td>
      <td> 
        <textarea name="gysxq" class="form" cols="45" rows="5"><%=gysxq%></textarea>
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td align="left">状态:</td>
      <td> 是否在首页显示本分类 
        <input type="checkbox" name="ifgystj" value="y" <%=isifgystj%>>
      </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td colspan="2" align="center"> 
        <input type="submit" name="Submit22" value="修 改">
        <input type="hidden" name="modigys" value=<%=request("modigys")%>><input type="hidden" name="id" value=<%=id%>>
        <input type="hidden" name="action" value="modi">
      </td>
    </tr>
  </form>
</table>
<%end if%>
<br>
<!--#include file="COPYRIGHT.ASP" -->
</body>
</html>