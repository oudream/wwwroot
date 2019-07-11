<!--#include file="SHEI45FECSA.ASP" -->
<%on error resume next%>
<%
'开始删除操作
if request("del")="true" then
conn.execute "delete * from myfaq where id=" & request("id")
url="viewfaq.asp?page=" & request("page")
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
    <td height="27">.:: 您可以在这里进行FAQ相关操作</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50"> 
      <form name="form1" method="post" action="VIEWFAQ.ASP">
        <div align="center">
          <select name="newstype">
            <option value="">Select ... </option>
            <option value="域名">域名 </option>
            <option value="虚拟主机">虚拟主机 </option>
            <option value="企业邮箱">企业邮箱 </option>
            <option value="网站建设">网站建设 </option>
            <option value="套餐">套餐 </option>
            <option value="租用托管">租用托管 </option>
            <option value="网站推广">网站推广 </option>
            <option value="域名知识">域名知识 </option>
            <option value="虚拟主机知识">虚拟主机知识 </option>
            <option value="企业邮箱知识">企业邮箱知识 </option>
            <option value="网站建设知识">网站建设知识 </option>
            <option value="套餐知识">套餐知识 </option>
            <option value="租用托管知识">租用托管知识 </option>
            <option value="网站推广知识">网站推广知识 </option>
          </select>
          &nbsp;&nbsp;&nbsp;<input type="submit" name="Submit" value="Submit">
        </div>
      </form></td>
  </tr>
</table>
<%
tnewstype=request("newstype")
if tnewstype<>"" then session("newstype")=tnewstype
sql="select * from myfaq where newstype='"&session("newstype")&"' order by id desc"
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
<tr><td></td><td>            
<%
zj=1
for zi=1 to rs.pagecount
%> 
   <a href="VIEWFAQ.ASP?page=<%=zj%>">|&nbsp;<%=zj%>&nbsp;|</a>         
<%
if zj mod 12 =0 then response.Write("<br><br>")
zj=zj+1
next
%> 
</td></tr>
<tr>
<td align="center" colspan="10">当前所在页:<font color=red><%=pagecount%></font>&nbsp;&nbsp;所在的项目：<font color=red><%=session("newstype")%></font></td>
</tr>
</table> <br>

<%
do while not rs.eof 
%>
     
<table border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" width="600" align="center">
  <tr bgcolor="#FFCC33"> 
    <td height="20" width="95" bgcolor="#FFCC33">&nbsp;FAQ/BASE标题：</td>
    <td width="194">&nbsp;<%=rs("subject")%></td>
    <td width="85">&nbsp;操作/日期：</td>
    <td width="205">&nbsp;<%=rs("idate")%> &nbsp;&nbsp;[<a href="VIEWFAQ.ASP?del=true&id=<%=rs("id")%>&page=<%=cstr(pagecount)%>">删除</a>] 
      &nbsp;&nbsp;[<a href="EDITFAQ.ASP?id=<%=rs("id")%>">编辑</a>] 
    </td>
  </tr>
  <tr bgcolor="#FFFF99"> 
    <td height="20" width="95" bgcolor="#FFFF99"> &nbsp;FAQ/BASE类型： </td>
    <td bgcolor="#FFFF99"> <font color="red"><%=rs("newstype")%></font> </td>
    <td bgcolor="#FFFF99">颜色：</td>
    <td bgcolor="<%=rs("color")%>"><font color="#FFFFFF"><%=rs("color")%></font>&nbsp;</td>
  </tr>
  <tr bgcolor="#FFFF99"> 
    <td height="45" bgcolor="#FFFF99">&nbsp;FAQ/BASE内容：</td>
    <td colspan="3" valign="top" bgcolor="#FFFF99"><table width="97%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td> <%  
message = replace(rs("message"),chr(13),"<br>")
message = replace(message,chr(32),"&nbsp;")
response.write message
%> </td>
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
        <a href="VIEWFAQ.ASP?page=<%=cstr(pagecount+1)%>"> 
        <font color="#000000">下一页</font></a> <% end if %> <% if rs.pagecount>1 and rs.pagecount=pagecount then %> 
        <a href="VIEWFAQ.ASP?page=<%=cstr(pagecount-1)%>"> 
        <font color="#000000">上一页</font></a> <%end if%> <% if pagecount<>1 and rs.pagecount<>pagecount then%> 
        <a href="VIEWFAQ.ASP?page=<%=cstr(pagecount-1)%>"> 
        <font color="#000000">上一页</font></a> <a href="VIEWFAQ.ASP?page=<%=cstr(pagecount+1)%>"> 
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
