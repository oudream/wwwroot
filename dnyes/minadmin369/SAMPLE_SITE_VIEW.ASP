<!--#include file="SHEI45FECSA.ASP" -->
<%on error resume next%>
<%
'开始删除操作
if request("del")="true" then
conn.execute "delete * from sample_site where id=" & request("id")
url="sample_site_view.asp?page=" & request("page")
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
    <td height="27">.:: 您可以在这里进行DELEDDOMAN相关操作</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>
<%
sql="select * from sample_site order by id desc"
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=6
rs.AbsolutePage=pagecount
%>
<table border="0" align="center" cellpadding="0" cellspacing="0">
<tr><td></td>            
<%
zj=1
for zi=1 to rs.pagecount
%> 
   <td width="50" align="center"><a href="sample_site_view.ASP?page=<%=zj%>">|&nbsp;<%=zj%>&nbsp;|</a></td>         
<%
zj=zj+1
next
%> 
</tr>
<tr>
<td align="center" colspan="10">当前所在页:<%=pagecount%></td>
</tr>
</table> <br>

<%
do while not rs.eof 
%>
     
<table border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" width="600" align="center">
  <tr bgcolor="#FFCC33"> 
    <td height="20" width="117" bgcolor="#FFCC33">网站名称：</td>
    <td width="172"><%=rs("sample_site_name")%>&nbsp;</td>
    <td width="85">网站地址：</td>
    <td width="205">&nbsp;<%=rs("addr")%> &nbsp;</td>
  </tr>
  <tr bgcolor="#FFFF99"> 
    <td bgcolor="#FFFF99">网站图片： </td>
    <td bgcolor="#FFFF99"><%=rs("img")%>&nbsp;</td>
    <td height="20" width="117" bgcolor="#FFFF99">&nbsp;操作</td>
    <td bgcolor="#FFFF99">&nbsp;[<a href="sample_site_view.ASP?del=true&id=<%=rs("id")%>&page=<%=cstr(pagecount)%>"  onClick="return confirm('确定删除此项吗？')">删除</a>] 
      &nbsp;&nbsp;[<a href="sample_site_edit.ASP?id=<%=rs("id")%>">编辑</a>] </td>
  </tr>
  <tr bgcolor="#FFFF99"> 
    <td height="45" bgcolor="#FFFF99">&nbsp;内容：</td>
    <td colspan="3" valign="top" bgcolor="#FFFF99"><table width="97%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td ><%=rs("memo")%>
		  </td>
        </tr>
      </table></td>
  </tr>
</table>
<br>
<%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
%>
      <br>
      
<table border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" width="65%" align="center">
  <tr bgcolor="#FFCC33"> 
    <td height="22"> 
      <div align="center">页次: <b><font color=red><%=pagecount%></font>/<%=rs.pagecount%></b> 
        <% if pagecount=1 and rs.pagecount<>pagecount and rs.pagecount<>0 then%> 
        <a href="sample_site_view.ASP?page=<%=cstr(pagecount+1)%>"> 
        <font color="#000000">下一页</font></a> <% end if %> <% if rs.pagecount>1 and rs.pagecount=pagecount then %> 
        <a href="sample_site_view.ASP?page=<%=cstr(pagecount-1)%>"> 
        <font color="#000000">上一页</font></a> <%end if%> <% if pagecount<>1 and rs.pagecount<>pagecount then%> 
        <a href="sample_site_view.ASP?page=<%=cstr(pagecount-1)%>"> 
        <font color="#000000">上一页</font></a> <a href="sample_site_view.ASP?page=<%=cstr(pagecount+1)%>"> 
        <font color="#000000">下一页</font></a> <% end if%> </div>
          </td>
        </tr>
      </table>
<%
rs.close
set rs=nothing
%>
<!--#include file="COPYRIGHT.ASP" -->
</body>
</html>
