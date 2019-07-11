<!--#include file="SHEI45FECSA.ASP" -->
<%on error resume next%>
<%
'开始删除操作
if request("action")="del" then
uid=request("uid")
viewcomp=request("viewcomp")
page=request("page")
conn.execute "delete * from orders where id=" & request("id")
url="userorderlog.asp?page=" & page & "&uid=" & uid & "&viewcomp=" & viewcomp
response.redirect url
end if
%>
<%
'开始完成操作
if request("action")="comp" then
uid=request("uid")
viewcomp=request("viewcomp")
username=request("username")
summoney=request("summoney")
jb=request("jb")
page=request("page")
'将orders表中的此订单设置为完成
conn.execute "update orders set comp=true where id="&request("id")
'将user表中此用户的可以使用折扣数增加，并增加其累计分数以在登陆时辨认级别
'取此用户级别应享受的折扣百分比,并计算出本次购物应享受的折扣进行累加
'注意：如果用户购物时使用了累计分，那么此处的购物总额为用户使用累计分后的总额，即实际支付额，所增加的累计分也为实际支付额的百分比
set rs=server.createobject("adodb.recordset") 
sql="select * from discount where leavel='"&jb&"'"
rs.open sql,conn,1,1
discount=rs("discount")
discount=summoney/100*discount
rs.close
set rs=nothing
'结束折扣计算
conn.execute "update user set discount=discount+'"&discount&"',sumjifen=sumjifen+'"&summoney&"' where uid='"&username&"'"
'取此用户的email以完成发信
set rs=server.createobject("adodb.recordset") 
sql="select * from uesr where uid='"&uid&"'"
rs.open sql,conn,1,1
email=rs("email")
rs.close
set rs=nothing
'向此用户发信，说明订单已被确认
fromemail=Application("y_it_fromemail")
toemail=email
mailsubject="来自" & Application("y_eshop_esname") & "的确认信，您的订单已被确认"  
mailbody="<font size=-1>您的订单已被确认</font>"

Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
objCDOMail.Importance=2
objCDOMail.BodyFormat=0
objCDOMail.MailFormat=0

objCDOMail.From=fromemail
objCDOMail.To=toemail
objCDOMail.Subject=mailsubject
objCDOMail.Body=mailbody
objCDOMail.Send
Set objCDOMail = Nothing
'返回操作页面
url="userorderlog.asp?page=" & page & "&uid=" & uid & "&viewcomp=" & viewcomp
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
    <td height="27">.:: 您可以在这里进行用户订单相关操作</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" width="90%" align="center">
  <tr bgcolor="#FFCC33"> 
    <td height="22">ID</td>
    <td bgcolor="#FFCC33">订单号</td>
    <td>支付金额</td>
    <td>支付方式</td>
    <td>订购日期</td>
    <td colspan="3">操作</td>
  </tr>
<%
uid=request("uid")
viewcomp=request("viewcomp")
sql="select * from orders where username<>'' "
if uid<>"" then sql=sql + " and username='"&uid&"'"
if viewcomp="comp" then sql=sql + " and comp=true"
if viewcomp="flase" then sql=sql + " and comp=flase"
sql=sql + " order by id desc"

set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1 

if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1   
end if

rs.pagesize=20
rs.AbsolutePage=pagecount
%> 
<% do while not rs.eof %> 
<tr bgcolor="#FFFF99"> 
    <td height="22" bgcolor="#FFFF99" width="11%"><%=rs("id")%></td>
    <td width="17%" bgcolor="#FFFF99"><%=rs("inBillNo")%></td>
    <td width="17%"><%=rs("summoney")%></td>
    <td width="17%" bgcolor="#FFFF99"><%=rs("paymenttype")%></td>
    <td width="17%"><%=rs("ordertime")%></td>
    <td width="7%">[<a href="ORDERSLOG.ASP?ordernumber=<%=rs("inBillNo")%>" target="_blank">详情</a>]</td>
    <td width="7%">[<a href="USERORDERLOG.ASP?action=del&id=<%=rs("id")%>&uid=<%=uid%>&viewcomp=<%=viewcomp%>&page=<%=cstr(pagecount)%>">删除</a>]</td>
    <td width="7%">[<a href="ORDERS.ASP?id=<%=rs("id")%>" target="_blank">状态</a>]</td>
  </tr>
<%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
%> 
<tr bgcolor="#FFFFCC"> 
<form action="USERORDERLOG.ASP" method="post">
      <td height="35" colspan="9"> 
        <div align="center"> 共 <b><%=rs.recordcount%></b> 张订单, 页次: <b><font color=red><%=pagecount%></font>/<%=rs.pagecount%></b>, 
          当前从第 <%
if pagecount<=1 then
response.write "<font color=red>1</font>"
else
response.write "<font color=red>" & pagecount*rs.pagesize-rs.pagesize+1 & "</font>"
end if
%> 张开始。 <% if pagecount=1 and rs.pagecount<>pagecount and rs.pagecount<>0 then%> 
          <a href="USERORDERLOG.ASP?uid=<%=uid%>&viewcomp=<%=viewcomp%>&page=<%=cstr(pagecount+1)%>"> 
          下一页</a> <% end if %> <% if rs.pagecount>1 and rs.pagecount=pagecount then %> 
          <a href="USERORDERLOG.ASP?uid=<%=uid%>&viewcomp=<%=viewcomp%>&page=<%=cstr(pagecount-1)%>"> 
          上一页</a> <%end if%> <% if pagecount<>1 and rs.pagecount<>pagecount then%> 
          <a href="USERORDERLOG.ASP?uid=<%=uid%>&viewcomp=<%=viewcomp%>&page=<%=cstr(pagecount-1)%>"> 
          上一页</a> <a href="USERORDERLOG.ASP?uid=<%=uid%>&viewcomp=<%=viewcomp%>&page=<%=cstr(pagecount+1)%>"> 下一页</a> 
          <% end if%>&nbsp; 直接到第 
          <select name="page">
            <%for i=1 to rs.pagecount%> 
            <option value="<%=i%>"><%=i%></option>
            <%next%> 
          </select>
          页 
          <input type="submit" name="go" value="Go">
          <input type="hidden" name="uid" value=<%=uid%>>
          <input type="hidden" name="viewcomp" value=<%=viewcomp%>>
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
