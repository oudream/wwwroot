<!--#include file="SHEI45FECSA.ASP" -->
<%on error resume next%>
<%
if request("action")="add" then
paymenttype=request("paymenttype")
paymentmessage=replace(request("paymentmessage"),">","&gt;")
paymentmessage=replace(paymentmessage,"<","&lt;")
idate=now()

conn.execute "insert into paydefault (paymenttype,paymentmessage,idate) values ('"&paymenttype&"','"&paymentmessage&"','"&idate&"')"
response.redirect "viewpayment.asp"
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

<table border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" width="63%" align="center">
  <form action="addpayment.asp" method="post" name="addpayment">
  <tr> 
    <td colspan="2" height="22" bgcolor="#FFCC33"><b>增加一个支付方式</b></td>
  </tr>
  <tr> 
    <td height="30" width="30%" bgcolor="#FFFF99">支付方式：</td>
    <td width="70%" bgcolor="#FFFF99">
      &nbsp;<input type=text name="paymenttype" size="32" maxlength="150" class="form">
    </td>
  </tr>
  <tr bgcolor="#FFCC33"> 
    <td height="30" valign="top"><br>
      反馈内容：</td>
    <td><br>&nbsp;
      <textarea cols="40" name="paymentmessage" rows="8"></textarea>
      <br>
      <br>
    </td>
  </tr>
  <tr bgcolor="#FFFFCC"> 
    <td colspan="2" height="35"> 
      <div align="center">
        <input type="submit" name="Submit" value="添加">
        <input type="hidden" name="action" value="add">
      </div>
    </td>
  </tr>
</form>
</table>
<!--#include file="COPYRIGHT.ASP" -->
</body>
</html>
