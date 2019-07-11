<!--#include file="SHEI45FECSA.ASP" -->
<%on error resume next%>
<%
'开始删除操作
if request("del")="true" then
uid=request("uid")
conn.execute "delete * from user where id=" & request("userid")
url="viewuser.asp?page=" & request("page") & "&uid=" & uid
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
    <td height="27">.:: 您可以在这里进行用户相关操作</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>

  <%
uid=request("uid")
if uid="" then
sql="select * from user order by id desc" 
else
sql="select * from user where uid like '%"&uid&"%' order by id desc" 
end if

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
<tr><td></td>            
<%
zj=1
for zi=1 to rs.pagecount
%> 
   <td width="50" align="center"><a href="VIEWUSER.ASP?page=<%=zj%>">|&nbsp;<%=zj%>&nbsp;|</a></td>         
<%
zj=zj+1
next
%> 
</tr>
<tr>
<td align="center" colspan="10">当前所在页:<%=pagecount%></td>
</tr>
</table> <br>
  
<table border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" width="97%" align="center">
  <tr> 
    <td height="22" bgcolor="#FFCC33"> ID </td>
    <td bgcolor="#FFCC33">用户名</td>
    <td bgcolor="#FFCC33">密码</td>
    <td bgcolor="#FFCC33">来自</td>
    <td bgcolor="#FFCC33">OICQ</td>
    <td bgcolor="#FFCC33">电话</td>
    <td bgcolor="#FFCC33">邮件地址</td>
    <td colspan="2" bgcolor="#FFCC33">操作</td>
  </tr>
<%
do while not rs.eof 
%> 
  <tr> 
    <td width="8%" height="22" bgcolor="#FFFF99"><%=rs("id")%></td>
    <td width="13%" bgcolor="#FFFF99"><a href="USERORDERLOG.ASP?uid=<%=rs("uid")%>"><%=rs("uid")%></a></td>
    <td width="10%" bgcolor="#FFFF99"><%=rs("pwd")%></td>
    <td width="15%" bgcolor="#FFFF99"><%=rs("city")%></td>
    <td width="10%" bgcolor="#FFFF99"><%=rs("oicq")%></td>
    <td width="10%" bgcolor="#FFFF99"><%=rs("tel")%></td>
    <td width="17%" bgcolor="#FFFF99"><a href="mailto:<%=rs("email")%>"><%=rs("email")%></a></td>
    <td width="11%" bgcolor="#FFFF99"><a href="EDITUSER.ASP?userid=<%=rs("id")%>">[查看|编辑]</a></td>
    <td width="6%" bgcolor="#FFFF99"><a href="VIEWUSER.ASP?del=true&userid=<%=rs("id")%>&page=<%=cstr(pagecount)%>&uid=<%=uid%>">[删除]</a> 
    </td>
  </tr>
<%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
%> 
<tr bgcolor="#FFFFCC"> 
<form action="VIEWUSER.ASP" method="post">
      <td height="35" colspan="9"> 
        <div align="center"> 共 <b><%=rs.recordcount%></b> 位用户, 页次: <b><font color=red><%=pagecount%></font>/<%=rs.pagecount%></b>, 
          当前从第 <%
if pagecount<=1 then
response.write "<font color=red>1</font>"
else
response.write "<font color=red>" & pagecount*rs.pagesize-rs.pagesize+1 & "</font>"
end if
%> 位开始。 <% if pagecount=1 and rs.pagecount<>pagecount and rs.pagecount<>0 then%> 
          <a href="VIEWUSER.ASP?uid=<%=uid%>&page=<%=cstr(pagecount+1)%>">下一页</a> 
          <% end if %> <% if rs.pagecount>1 and rs.pagecount=pagecount then %> 
          <a href="VIEWUSER.ASP?uid=<%=uid%>&page=<%=cstr(pagecount-1)%>"> 上一页</a> 
          <%end if%> <% if pagecount<>1 and rs.pagecount<>pagecount then%> <a href="VIEWUSER.ASP?uid=<%=uid%>&page=<%=cstr(pagecount-1)%>"> 
          上一页</a> <a href="VIEWUSER.ASP?uid=<%=uid%>&page=<%=cstr(pagecount+1)%>"> 
          下一页</a> <% end if%>&nbsp; 直接到第 
          <select name="page">
            <%for i=1 to rs.pagecount%> 
            <option value="<%=i%>"><%=i%></option>
            <%next%> 
          </select>
          页 
          <input type="submit" name="go" value="Go">
          <input type="hidden" name="uid" value=<%=uid%>>
        </div>
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
