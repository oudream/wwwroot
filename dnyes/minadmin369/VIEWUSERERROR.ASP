<!--#include file="SHEI45FECSA.ASP" -->
<%on error resume next%>
<%
'开始删除操作
if request("action")="del" then
id=request("id")
conn.execute "delete * from usererror where id=" & id
url="viewusererror.asp?page=" & request("page")
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
    <td height="27">.:: 您可以在这里进行用户投诉订单相关操作</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>

<table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" width="90%" align="center">
  <tr bgcolor="#FFCC33"> 
    <td width="10%">ID</td>
    <td width="25%">用户帐号</td>
    <td width="25%">订单号</td>
    <td width="25%">申诉日期</td>
    <td width="15%">操作</td>
  </tr>
<%
sql="select * from usererror order by id desc"
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=20
rs.AbsolutePage=pagecount

do while not rs.eof 
%>
  <tr bgcolor="#FFFF99"> 
    <td><%=rs("id")%></td>
    <td><%=rs("username")%></td>
    <td><%=rs("inBillNo")%></td>
    <td><%=rs("sdate")%></td>
    <td bgcolor="#FFFF99">[<a href="VIEWUSERERROR.ASP?action=del&id=<%=rs("id")%>&page=<%=cstr(pagecount)%>">删除</a>]  [<a href="ORDERSLOG.ASP?ordernumber=<%=rs("inBillNo")%>" target="_blank">详情</a>]</td>
  </tr>
<%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
%> 
<tr bgcolor="#FFFFCC"> 
<form action="VIEWUSERERROR.ASP" method="post">
      <td height="35" colspan="5"> 
        <div align="center"> 共 <b><%=rs.recordcount%></b> 张单, 页次: <b><font color=red><%=pagecount%></font>/<%=rs.pagecount%></b>, 
          当前从第 <%
if pagecount<=1 then
response.write "<font color=red>1</font>"
else
response.write "<font color=red>" & pagecount*rs.pagesize-rs.pagesize+1 & "</font>"
end if
%> 张开始。 <% if pagecount=1 and rs.pagecount<>pagecount and rs.pagecount<>0 then%> 
          <a href="VIEWUSERERROR.ASP?page=<%=cstr(pagecount+1)%>"><font color="#000000">下一页</font></a> 
          <% end if %> <% if rs.pagecount>1 and rs.pagecount=pagecount then %> 
          <a href="VIEWUSERERROR.ASP?&page=<%=cstr(pagecount-1)%>"> <font color="#000000">上一页</font></a> 
          <%end if%> <% if pagecount<>1 and rs.pagecount<>pagecount then%> <a href="VIEWUSERERROR.ASP?page=<%=cstr(pagecount-1)%>"> 
          <font color="#000000">上一页</font></a> <a href="VIEWUSERERROR.ASP?page=<%=cstr(pagecount+1)%>"> 
          <font color="#000000">下一页</font></a> <% end if%>&nbsp; 直接到第 
          <select name="page">
            <%for i=1 to rs.pagecount%> 
            <option value="<%=i%>"><%=i%></option>
            <%next%> 
          </select>
          页 
          <input type="submit" name="go" value="Go"></div>
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
