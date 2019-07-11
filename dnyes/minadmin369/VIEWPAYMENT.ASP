<!--#include file="SHEI45FECSA.ASP" -->
<%on error resume next%>
<%
'开始删除操作
if request("del")="true" then
conn.execute "delete * from paydefault where id=" & request("id")
url="viewpayment.asp?page=" & request("page")
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
    <td height="27">.:: 您可以在这里进行支付方式相关操作</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>
<%
sql= "select * from paydefault order by id desc"
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 

if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=20
rs.AbsolutePage=pagecount

do while not rs.eof 
%>
      
<table border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" width="65%" align="center">
  <tr bgcolor="#FFCC33"> 
    <td height="22" width="15%" bgcolor="#FFCC33">&nbsp;支付方式：</td>
          
    <td width="35%">&nbsp;<%=rs("paymenttype")%></td>
          
    <td width="15%">&nbsp;操作/日期：</td>
          
    <td width="35%">&nbsp;<%=rs("idate")%>
&nbsp;&nbsp;[<a href="viewpayment.asp?del=true&id=<%=rs("id")%>&page=<%=cstr(pagecount)%>">删除</a>]
</td>
        </tr>
        
  <tr bgcolor="#FFFF99"> 
    <td height="45" width="15%" valign="top" bgcolor="#FFFF99"><br>&nbsp;反馈内容：</td>
    <td colspan="3" valign="top" bgcolor="#FFFF99"> <br>
      <table width="97%" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr>
                <td> 
<%
    paymentmessage = rs("paymentmessage")
	paymentmessage = replace(paymentmessage,chr(13),"<br>")
    paymentmessage = replace(paymentmessage,chr(32),"&nbsp;")
	response.write paymentmessage
%>
</td>
              </tr>
            </table><br>
    </td>
        </tr>
      </table><br>
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
        <a href="viewpayment.asp?page=<%=cstr(pagecount+1)%>"> 
        <font color="#000000">下一页</font></a> <% end if %> <% if rs.pagecount>1 and rs.pagecount=pagecount then %> 
        <a href="viewpayment.asp?page=<%=cstr(pagecount-1)%>"> 
        <font color="#000000">上一页</font></a> <%end if%> <% if pagecount<>1 and rs.pagecount<>pagecount then%> 
        <a href="viewpayment.asp?page=<%=cstr(pagecount-1)%>"> 
        <font color="#000000">上一页</font></a> <a href="viewpayment.asp?page=<%=cstr(pagecount+1)%>"> 
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
