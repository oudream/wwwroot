<!--#include file="SHEI45FECSA.ASP" -->
<%
'修改公告
if request("action")="modi" then

deleddate=request("deleddate")
color=request("color")
if color="" then color=" "
deleddomain=request("deleddomain")
deledtable=cstr(datevalue(deleddate))+"失效的3位国际域名"

conn.execute "update deleddomain set deledtable='"&deledtable&"',deleddomain='"&deleddomain&"',deleddate='"&deleddate&"',color='"&color&"' where id=" & request("id")
response.redirect "viewdeleddomain.asp"
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
    <td height="27">.:: 您可以在这里进行站内公告相关操作</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>
<%
sql="select * from deleddomain where id=" & request("id")
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
%> 
<table border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" width="600" align="center">
  <form action="EDITDELEDDOMAIN.ASP" method="post" name="news">
    <tr bgcolor="#FFCC33"> 
      <td height="33" width="28%">&nbsp;deleddomain标题：</td>
      <td colspan="3"> &nbsp; <input maxlength=45 name=deledtable size=60 class="form" value="<%=rs("deledtable")%>">
        -->>ID:<%=rs("id")%> </td>
    </tr>
    <tr bgcolor="#FFCC33"> 
      <td height="36" width="28%">deleddomain日期：</td>
      <td width="22%"><input maxlength=45 name=deleddate size=15 class="form" value="<%=rs("deleddate")%>"> 
      </td>
      <td width="14%">COLOR:</td>
      <td width="36%" bgcolor="<%=rs("color")%>"><input maxlength=6 name=color id=color size=10 class="form" value="<%=rs("color")%>"></td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td colspan="4">&nbsp;</td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td height="45">&nbsp;deleddomain内容：</td>
      <td height="45" colspan="3"><textarea name="deleddomain" cols="60" class="form" rows="8">
<%=rs("deleddomain")%>
</textarea></td>
    </tr>
    <tr align="center" bgcolor="#FFFF99"> 
      <td height="45" colspan="4"><input type="submit" name="Submit" value="确 定"> 
        &nbsp;&nbsp; <input type="reset" name="Submit" value="取 消"> &nbsp;&nbsp; 
        <input type="button" name="Submit" value="后 退" onClick="history.go( -1 );return true;"> 
        <input type="hidden" name="action" value="modi"> <input type="hidden" name="ID" value="<%=request("id")%>"> 
      </td>
    </tr>
  </form>
</table>
<%
rs.close
set rs=nothing
%> 
<!--#include file="COPYRIGHT.ASP" -->
</body>
</html>
