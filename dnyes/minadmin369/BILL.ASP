<!--#include file="SHEI45FECSA.ASP" -->
<%on error resume next%>
<%
'修改广告
if request("action")="modi" then

subject=request("subject")
newstype=request("newstype")
message=request("message")
conn.execute "update message set subject='"&subject&"',message='"&message&"' where newstype='bill'"
response.redirect "bill.asp"
end if 
%>
<html>
<head>
<title>黄页系统管理中心</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../inc/southdns.css" type="text/css">
</head>
<body bgcolor="#FFFFCC" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" vspace="0" hspace="0">
  <tr bgcolor="#FFCC33"> 
    <td height="27">.:: 您可以在这里进行站内广告相关操作</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>
<%
sql="select * from message where newstype='bill'"
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
%> 
<table border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" width="65%" align="center">
<form action="bill.asp" method="post" name="bill">
  <tr bgcolor="#FFCC33"> 
    <td height="22" width="25%">&nbsp;广告标题：</td>
      <td> &nbsp;
        <input maxlength=45 name=subject size=40 class="form" value="<%=rs("subject")%>">
      </td>
  </tr>
  <tr bgcolor="#FFFF99"> 
    <td height="22" valign="top"><br>
      &nbsp;广告内容：</td>
    <td><br>
        &nbsp; 
        <textarea name="message" cols="50" class="form" rows="8">
<%=rs("message")%>
</textarea>
        <br>
        <br>
      </td>
  </tr>
  <tr bgcolor="#FFFF99"> 
    <td height="45" colspan="2"> 
      <div align="center"> 
        <input type="submit" name="Submit" value="确 定">
        <input type="reset" name="Submit" value="取 消">
        <input type="hidden" name="action" value="modi">
      </div>
    </td>
  </tr>
</form>
</table>
<%
rs.close
set rs=nothing
%> 
</body>
</html>
