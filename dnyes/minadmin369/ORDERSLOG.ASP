<!--#include file="SHEI45FECSA.ASP" -->
<%on error resume next%>
<%
sql="select * from orders where inBillNo='"&request("ordernumber")&"'"
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
if rs.eof or rs.bof then
rs.close
set rs=nothing
response.write"<SCRIPT language=JavaScript>alert('对不起，不存在此订单,请核实！');"
response.write"javascript:history.go(-1)</SCRIPT>"
response.end
end if
username=rs("username")
sqlt="select * from user where uid='"&username&"'"
set rst=server.createobject("ADODB.Recordset")
rst.open sqlt,conn,1,1
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
  <tr> 
    <td bgcolor="#FFCC33" height="22" width="25%">订单号：</td>
    <td bgcolor="#FFCC33" width="25%"><%=rs("inBillNo")%></td>
    <td bgcolor="#FFCC33" width="25%">订单总金额：</td>
    <td bgcolor="#FFCC33" width="25%"><%=rs("summoney")%> RMB       [<a href="ORDERS.ASP?id=<%=rs("id")%>" target="_blank">状态</a>]</td>
  </tr>
  <tr> 
    <td bgcolor="#FFFF99" height="22">订购帐号：</td>
    <td bgcolor="#FFFF99"><a href="EDITUSER.ASP?userid=<%=rst("id")%>"><b><%=rst("uid")%></b></a></td>
    <td bgcolor="#FFFF99">订购时间：</td>
    <td bgcolor="#FFFF99"><%=rs("ordertime")%></td>
  </tr>
  <tr bgcolor="#FFCC33"> 
    <td height="22">订购人电话：</td>
    <td><%=rst("tel")%></td>
    <td>订购人Email：</td>
    <td><a href="mailto:<%=rst("email")%>"><%=rst("email")%></a></td>
  </tr><tr bgcolor="#FFFF99"> 
    <td height="22">支付方式：</td>
    <td><%=rs("paymenttype")%></td>
    <td>备用帐号：</td>
    <td><%=rs("glzh2")%></td>
  </tr>
  <tr bgcolor="#FFCC33"> 
    <td height="22">此订单状态：</td>
    <td><%if rs("ifcomp")=true then
response.write "已经开通"
else
if rs("ifpay")=true then
response.write "付款已到，尚未开通。"
else
response.write "尚未付款"
end if
end if
%></td>
    <td>开通时间：</td>
    <td><%=rs("kttime")%></td>
  </tr><%if rs("iffalse")=true then%>
          <tr bgcolor="#FFFF99">
            <td height="22"><font color="red"><b>此订单错误信息：</b></font></td>
            <td height="22" colspan="3">
<%
errmsg=replace(rs("errmsg"),chr(13),"<br>")
errmsg=replace(errmsg,chr(32),"&nbsp;")
response.write errmsg
%>
</td>
          </tr>
<%end if%>
</table><br>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" width="90%" align="center">
    
  <tr bgcolor="#FFCC33"> 
    <td align="left" height="22" colspan="5"> 
      <div align="left"><b>此定单详细情况</b></div>
      </td>
    </tr>
  <%
xiangxi=rs("xiangxi")
xiangxi1=rs("xiangxi1")
xxgys=rs("xxgys")
glzh=rs("glzh")
glma=rs("glma")

xiangxit=split(xiangxi,",")
xiangxi1t=split(xiangxi1,",")
xxgys=split(xxgys,",")
glzhs=split(glzh,",")
glmas=split(glma,",")
for i=0 to ubound(xiangxit)
%> 
    <td width="20%" height="22" bgcolor="#FFFF99"><%=xiangxit(i)%></td>
            <td width="20%" bgcolor="#FFFF99">供应商：<%=xxgys(i)%></td><td width="20%" bgcolor="#FFFF99">管理帐号：<%=glzhs(i)%></td><td width="22%" bgcolor="#FFFF99">初始密码：<%=glmas(i)%></td>
            <td width="18%" bgcolor="#FFFF99">购买年限：<%=xiangxi1t(i)%>年</td>
    </tr>
<%
next
%>
  <tr bgcolor="#FFFF99"> 
    <td colspan="5">本定单共以上商品</td>
  </tr>
</table>
<%
rs.close
set rs=nothing
rst.close
set rst=nothing
%>
<!--#include file="COPYRIGHT.ASP" -->
</body>
</html>
