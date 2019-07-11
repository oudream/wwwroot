<!--#include file="SHEI45FECSA.ASP" -->
<%on error resume next%>
<%
'开始删除操作
if request("del")="true" then
id=request("id")
conn.execute "delete * from subs where id=" & id
url="viewsubs.asp?page=" & request("page") & "&subs=" & subs & "&area=" & area
response.redirect url
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
    <td height="27">.:: 您可以在这里进行商品相关操作</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="50%" height="103" align="center"><table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" align="center" width="100%">
        <tr bgcolor="#FFCC33"> 
          <td colspan="2">请选择您要批量浏览的商品的条件</td>
        </tr>
        <tr bgcolor="#FFFFCC"> 
          <td colspan="2">由于查询全部商品过慢，所以请选择商品所在分类 </td>
        </tr>
        <form name="searchsubs" method="post" action="viewsubs.asp">
          <tr bgcolor="#FFFF99"> 
            <td align="right" width="35%">商品分类:</td>
            <td> <select name="area">
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
              </select> </td>
          </tr>
          <tr bgcolor="#FFFF99"> 
            <td colspan="2" align="center"> <input type="submit" name="Submit" value="确 定"> 
            </td>
          </tr>
        </form>
      </table></td>
    <td align="center"><table width="101%" height="18%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#000000">
        <tr bgcolor="#FFCC33"> 
          <td colspan="3">&nbsp;</td>
        </tr>
        <form name="searchuser" method="post" action="viewsubs.asp">
          <tr bgcolor="#FFFF99"> 
            <td width="31%" align="center">商品名:</td>
            <td width="29%" align="center">ID</td>
            <td width="40%" align="center">Price &gt;=</td>
          </tr>
          <tr bgcolor="#FFFF99"> 
            <td height="28" align="center"><input type="text" name="subsname2" class="form"> 
            </td>
            <td align="center"><input name="sid" type="text" class="form" id="sid" size="8"></td>
            <td align="center"><input name="price" type="text" class="form" id="price" size="12"></td>
          </tr>
          <tr bgcolor="#FFFF99"> 
            <td height="26" colspan="3" align="center"><input type="submit" name="Submit2" value="确 定"></td>
          </tr>
        </form>
      </table>
</td>
  </tr>
</table>
<br>

<%
subsname=request("subsname")
area=request("area")
sid=request("sid")
if sid<>"" then session("price")=""
price=request("price")
if price<>"" then session("price")=price
sql="select * from subs where subsname<>'' "
if subsname<>"" then sql=sql + " and subsname like '%"&subsname&"%'"
if area<>"" then sql=sql + " and area='"&area&"'"
if session("price")<>"" then sql=sql + " and price>="&clng(session("price"))
if sid<>"" then sql=sql + " and id="&clng(sid)
sql=sql + " order by id desc"

Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=20
rs.AbsolutePage=pagecount
%>
<table border="0" align="center" cellpadding="0" cellspacing="0">
<tr><td></td><td>            
<%
zj=1
for zi=1 to rs.pagecount
%> 
   <a href="viewsubs.asp?subsname=<%=subsname%>&area=<%=area%>&page=<%=zj%>">|&nbsp;<%=zj%>&nbsp;|</a>         
<%
if zj mod 12 =0 then response.Write("<br><br>")
zj=zj+1
next
%> 
</td></tr>
<tr>
<td align="center" colspan="10">当前所在页:<font color=red><%=pagecount%></font>&nbsp;&nbsp;所在的项目：<font color=red><%=request("area")%></font></td>
</tr>
</table>
 <br>
<table border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" width="97%" align="center">
  <tr> 
    <td bgcolor="#FFCC33" height="22">ID</td>
    <td bgcolor="#FFCC33">商品分类</td>
    <td bgcolor="#FFCC33">商品名称</td>
    <td bgcolor="#FFCC33">供应商</td>
    <td bgcolor="#FFCC33">价格</td>
    <td bgcolor="#FFCC33">是否推荐</td>
    <td bgcolor="#FFCC33" colspan="2">操作</td>
  </tr>

<%
do while not rs.eof 
%> 
  <tr> 
    <td height="22" bgcolor="#FFFF99" width="8%"><%=rs("id")%></td>
    <td bgcolor="#FFFF99" width="10%"><%=rs("area")%></td>
    <td bgcolor="#FFFF99" width="24%"><%=rs("subsname")%></td>
    <td bgcolor="#FFFF99" width="18%"><%=rs("gys")%></td>
    <td bgcolor="#FFFF99" width="15%" align="right"><%=rs("price")%> 元</td>
<%iftj="" 
if rs("iftj")=true then iftj="已推荐" 
%>
    <td bgcolor="#FFFF99" width="9%"><%=iftj%></td>
    <td bgcolor="#FFFF99" width="10%">
	<%if area="网站建设" then
	response.Write("<a href=editsubs-web.asp?idd="&rs("idd")&">[查看|编辑]</a>")
	else
	response.Write("<a href=editsubs.asp?idd="&rs("idd")&">[查看|编辑]</a>")
	end if
	%>
	</td>
    <td bgcolor="#FFFF99" width="6%">
<a href="viewsubs.asp?del=true&id=<%=rs("id")%>&subsname=<%=subsname%>&area=<%=area%>&page=<%=cstr(pagecount)%>">[删除]</a> 
    </td>
  </tr>
<%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
%> 
<tr bgcolor="#FFFFCC"> 
<form action="viewsubs.asp" method="post">
      <td height="35" colspan="9"> 
        <div align="center"> 共 <b><%=rs.recordcount%></b> 个商品, 页次: <b><font color=red><%=pagecount%></font>/<%=rs.pagecount%></b>, 
          当前从第 <%
if pagecount<=1 then
response.write "<font color=red>1</font>"
else
response.write "<font color=red>" & pagecount*rs.pagesize-rs.pagesize+1 & "</font>"
end if
%> 个开始。 <% if pagecount=1 and rs.pagecount<>pagecount and rs.pagecount<>0 then%> 
<a href="viewsubs.asp?subsname=<%=subsname%>&area=<%=area%>&page=<%=cstr(pagecount+1)%>">下一页</a> 
          <% end if %> <% if rs.pagecount>1 and rs.pagecount=pagecount then %> 
<a href="viewsubs.asp?subsname=<%=subsname%>&area=<%=area%>&page=<%=cstr(pagecount-1)%>"> 上一页</a> 
          <%end if%> <% if pagecount<>1 and rs.pagecount<>pagecount then%> 
<a href="viewsubs.asp?subsname=<%=subsname%>&area=<%=area%>&page=<%=cstr(pagecount-1)%>"> 上一页</a> 
<a href="viewsubs.asp?subsname=<%=subsname%>&area=<%=area%>&page=<%=cstr(pagecount+1)%>"> 下一页</a> <% end if%>&nbsp; 直接到第 
          <select name="page">
            <%for i=1 to rs.pagecount%> 
            <option value="<%=i%>"><%=i%></option>
            <%next%> 
          </select>
          页 
          <input type="submit" name="go" value="Go">
          <input type="hidden" name="subsname" value=<%=subsname%>>
          <input type="hidden" name="area" value=<%=area%>></div>
</td></form>
</tr>
</table>
<%
rs.close
set rs=nothing
%>
<!--#include file="COPYRIGHT.ASP" -->
</body>
</html>
