<!--#include file="SHEI45FECSA.ASP" -->
<%on error resume next%>
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







    <form action="ORDERSELECT.ASP" method="post">

  <table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" width="90%" align="center">
    <tr bgcolor="#FFCC33"> 
      <td height="22" align="center">UID</td>
      <td align="center" bgcolor="#FFCC33">订单号</td>
      <td align="center">订购日期</td>
      <td>支付方式</td>
      <td align="center">产品类型</td>
      <td width="30%"> <input name="selifpay" type="checkbox" id="selifpay" value="y" <%=ifpay%>>
        已付款 
        <input name="selifcomp" type="checkbox" id="selifcomp" value="y" <%=ifcomp%>>
        已开通 </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td height="22" bgcolor="#FFFF99" width="10%"><input name="seluid" type="text" id="seluid" size="12"></td>
      <td width="10%" bgcolor="#FFFF99"><input name="selinbillno" type="text" id="selinbillno" size="12"></td>
      <td width="28%"><input name="selyear" type="text" id="selyear" size="6">
        y 
        <input name="selmonth" type="text" id="selmonth" size="4">
        m 
        <input name="selday" type="text" id="selday" size="4">
        d </td>
      <td width="11%" bgcolor="#FFFF99"> <select name="selpaymenttype" id="selpaymenttype">
          <option value="" selected>支付方式</option>
          <%
sqlp="select paymenttype from paydefault"
set rsp=server.createobject("ADODB.Recordset")
rsp.open sqlp,conn,1,1
while not rsp.eof
%>
          <option value="<%=rsp("paymenttype")%>"><%=rsp("paymenttype")%></option>
          <%
rsp.movenext
wend
rsp.close
set rsp=nothing
%>
        </select> </td>
      <td width="11%"><select name="selsubid" id="selsubid">
          <option value="" selected>产品ID</option>
<%
sqlp="select id from subs"
set rsp=server.createobject("ADODB.Recordset")
rsp.open sqlp,conn,1,1
while not rsp.eof
%>
          <option value="<%=rsp("id")%>"><%=rsp("id")%></option>
          <%
rsp.movenext
wend
rsp.close
set rsp=nothing
%>

        </select></td>
      <td><input name="seliffalse" type="checkbox" id="seliffalse" value="y" <%=iffalse%>>
        订单错误 
        <input name="selifdeal" type="checkbox" id="selifdeal" value="y" <%=ifdeal%>>
        订单处理中 </td>
    </tr>
    <tr bgcolor="#FFFFCC"> 
      <td height="35" colspan="7"> <div align="center">
          <input name="selecting" type="submit" id="selecting" value="Select...">
        </div></td>
    </tr>
  </table>
    </form>














<table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" width="90%" align="center">
  <tr bgcolor="#FFCC33"> 
    <td height="22">ID</td>
    <td bgcolor="#FFCC33">订单号</td>
    <td>支付金额</td>
    <td>支付方式</td>
    <td>订购日期</td>
    <td colspan="3">操作 </td>
  </tr>
<%
seluid=request("seluid")
selinbillno=request("selinbillno")

selyear=request("seldate")
selmonth=request("selmonth")
selday=request("selday")
seldate="%"
if trim(selyear)<>"" then seldate=seldate + selyear + "[-]" 
if trim(selmonth)<>"" then seldate=seldate + selmonth + "[-]"  + "%"
if trim(selday)<>"" then seldate=seldate + selday + "%"
selpaymenttype=request("selpaymenttype")
selsubid=request("selsubid")
selifpay=request("selifpay")
selifcomp=request("selifcomp")
seliffalse=request("seliffalse")
selifdeal=request("selifdeal")
uid=request("uid")
viewcomp=request("viewcomp")

sql="select * from orders where username<>'' "

if trim(seluid)<>"" then sql=sql + " and username='"&seluid&"'"
if trim(selinbillno)<>"" then sql=sql + " and inbillno='"&selinbillno&"'"
if trim(seldate)<>"" then sql=sql + " and ordertime like '"&seldate&"'"
if trim(selpaymenttype)<>"" then sql=sql + " and paymenttype='"&selpaymenttype&"'"
if trim(selsubid)<>"" then sql=sql + " and subid='"&selsubid&"'"
if trim(selifpay)="y" then sql=sql + " and ifpay=true"
if trim(selifcomp)="y" then sql=sql + " and ifcomp=true"
if trim(seliffalse)="y" then sql=sql + " and iffalse=true"
if trim(selifdeal)="y" then sql=sql + " and ifdeal=true"
if uid<>"" then sql=sql + " and username='"&uid&"'"
if viewcomp="comp" then sql=sql + " and comp=true"
if viewcomp="flase" then sql=sql + " and comp=flase"
%> 
<%
sql=sql + " order by id desc"

set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1 

if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1   
end if

rs.pagesize=200
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
    <td width="7%">[<a href="ORDERS.ASP?id=<%=rs("id")%>" target="_blank">状态</a>]</td>
  </tr>
<%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
%> 
<tr bgcolor="#FFFFCC"> 
<form action="ORDERSELECT.ASP" method="post">
      <td height="35" colspan="9"> 
        <div align="center"> 共 <b><%=rs.recordcount%></b> 张订单, 页次: <b><font color=red><%=pagecount%></font>/<%=rs.pagecount%></b>, 
          当前从第 <%
if pagecount<=1 then
response.write "<font color=red>1</font>"
else
response.write "<font color=red>" & pagecount*rs.pagesize-rs.pagesize+1 & "</font>"
end if
%> 张开始。 <% if pagecount=1 and rs.pagecount<>pagecount and rs.pagecount<>0 then%> 
          <a href="ORDERSELECT.ASP?uid=<%=uid%>&viewcomp=<%=viewcomp%>&page=<%=cstr(pagecount+1)%>"> 
          下一页</a> <% end if %> <% if rs.pagecount>1 and rs.pagecount=pagecount then %> 
          <a href="ORDERSELECT.ASP?uid=<%=uid%>&viewcomp=<%=viewcomp%>&page=<%=cstr(pagecount-1)%>"> 
          上一页</a> <%end if%> <% if pagecount<>1 and rs.pagecount<>pagecount then%> 
          <a href="ORDERSELECT.ASP?uid=<%=uid%>&viewcomp=<%=viewcomp%>&page=<%=cstr(pagecount-1)%>"> 
          上一页</a> <a href="ORDERSELECT.ASP?uid=<%=uid%>&viewcomp=<%=viewcomp%>&page=<%=cstr(pagecount+1)%>"> 下一页</a> 
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
