<!--#include file="SHEI45FECSA.ASP" -->
<%on error resume next%>

<%
'开始完成操作
if request("action")="modi" then
id=request("id")
errmsg=replace(request("errmsg"),">","&gt;")
errmsg=replace(errmsg,"<","&lt;")
errmsg=replace(errmsg,"'","’")
conn.execute "update orders set errmsg='"&errmsg&"' where id="&id

if request("ifcomp")="y" then
conn.execute "update orders set ifcomp=true where id="&id
conn.execute "update orders set kttime=now() where id="&id
'可以发信

end if
if request("ifcomp")="" then conn.execute "update orders set ifcomp=false where id="&id
if request("ifpay")="y" then conn.execute "update orders set ifpay=true where id="&id
if request("ifpay")="" then conn.execute "update orders set ifpay=false where id="&id
if request("iffalse")="y" then conn.execute "update orders set iffalse=true where id="&id
if request("iffalse")="" then conn.execute "update orders set iffalse=false where id="&id
if request("ifdeal")="y" then conn.execute "update orders set ifdeal=true where id=" & id
if request("ifdeal")="" then conn.execute "update orders set ifdeal=false where id=" & id


'返回操作页面
url="orders.asp?id="&id
response.write "订单资料已修改成功，本页面2秒后返回."
response.write "<meta http-equiv=refresh content=2;url=" & url & ">"
response.end
end if
%>
<%
id=request("id")
set rs=server.createobject("adodb.recordset") 
sql="select * from orders where id=" & id
rs.open sql,conn,1,1
ifpay=""
ifcomp=""
iffalse=""
ifdeal=""
if rs("ifpay")=True then ifpay="checked"
if rs("ifcomp")=True then ifcomp="checked"
if rs("iffalse")=True then iffalse="checked"
if rs("ifdeal")=true then ifdeal="checked"
username=rs("username")


if rs("errmsg")="" or rs("errmsg")="无" then 
errmsg="此订单并非错误订单"
else
errmsg=rs("errmsg")
end if

rs.close
set rs=nothing
%>
<html>
<head>
<title>商城系统管理中心</title>
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

<table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" width="450" align="center">
  <form action="ORDERS.ASP" method="post" name="orders">
  <tr> 
    <td bgcolor="#FFCC33" width="70">定单状态：</td>
    <td bgcolor="#FFCC33">
        <input type="checkbox" name="ifpay" value="y" <%=ifpay%>>
        已付款 
        <input type="checkbox" name="ifcomp" value="y" <%=ifcomp%>>
        已开通 
        <input type="checkbox" name="iffalse" value="y" <%=iffalse%>>
        订单错误 
		<input type="checkbox" name="ifdeal" value="y" <%=ifdeal%>>
		订单处理中
		</td>
  </tr>
  <tr> 
      <td bgcolor="#FFFF99" valign="top"><br>
        错误信息：</td>
    <td bgcolor="#FFFF99">
        <textarea name="errmsg" rows="6" cols="40"><%=errmsg%></textarea>
      </td>
  </tr>
  <tr bgcolor="#FFFFCC"> 
      <td colspan="2"> 
        <div align="center">
          <input type="submit" name="Submit" value="修改">
          &nbsp;&nbsp; 
          <input type="reset" name="Submit2" value="重填">
          <input type="hidden" name="id" value="<%=id%>">
          <input type="hidden" name="action" value="modi"><input type="hidden" name="username" value="<%=username%>">
        </div>
      </td>
  </tr>
</form>
</table>
</body>
</html>
