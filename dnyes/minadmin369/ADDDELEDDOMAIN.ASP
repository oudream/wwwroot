<!--#include file="SHEI45FECSA.ASP" -->
<%
'添加公告
if request("action")="add" then

deleddate=request("deleddate")
sql="select * from deleddomain where deleddate='" &request("deleddate")&"'"
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
if not (rs.bof or rs.eof) then 
response.Write("<table width=100% height=100% border=0 align=center cellpadding=0 cellspacing=0><tr><td align=center>此日期已经存在，请选择编辑<br><br></td></tr></table>")
response.End()
rs.close
set rs=nothing
end if
color=request("color")
if color="" then color=" "
deleddomain=request("deleddomain")
deledtable=deleddate+request("subject")
conn.execute "insert into deleddomain (deledtable,deleddomain,deleddate,color) values ('"&deledtable&"','"&deleddomain&"','"&deleddate&"','"&color&"')"

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
    <td height="27">.:: 您可以在这里进行DELEDDOMAIN相关操作</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
</table>
<table border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" width="65%" align="center">
  <form action="ADDDELEDDOMAIN.ASP" method="post" name="addnews">
    <tr bgcolor="#FFCC33"> 
      <td height="22" width="16%">添加标题：</td>
      <td colspan="3"><input maxlength=45 name="subject" size=30 class="form" value="失效的3位国际域名"></td>
    </tr>
    <tr bgcolor="#FFCC33"> 
      <td width="16%" height="22">添加日期：</td>
      <td width="39%"><input maxlength=45 name="deleddate" size=30 class="form" value="<%=date%>"> 
      </td>
      <td width="19%">DELEDDOMAIN 颜色：</td>
      <td width="26%"><input  name="color" size="12" maxlength=6 class="form"></td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td colspan="4"> <div align="center"></div></td>
    </tr>
    <tr bgcolor="#FFFFCC"> 
      <td height="45">&nbsp;内容：</td>
      <td height="45" colspan="3"><textarea name="deleddomain" cols="60" rows="8" class="form" id="deleddomain">
</textarea></td>
    </tr>
    <tr bgcolor="#FFFFCC"> 
      <td height="45" colspan="4" align="center"> <input type="submit" name="Submit" value="确 定"> 
        &nbsp;&nbsp; <input type="reset" name="Submit" value="取 消"> &nbsp;&nbsp; 
        <input type="button" name="Submit" value="后 退" onClick="history.go( -1 );return true;"> 
        <input type="hidden" name="action" value="add"> </td>
    </tr>
  </form>
</table>
<!--#include file="COPYRIGHT.ASP" -->
</body>
</html>
