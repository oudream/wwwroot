<!--#include file="SHEI45FECSA.ASP" -->
<%on error resume next%>
<%
'修改公告
if request("action")="modi" then
color=request("color")
if trim(color)="" then color=" "
indate=request("indate")
if trim(indate)="" then indate=" "
subject=request("subject")
addr=request("addr")
conn.execute "update document set subject='"&subject&"',addr='"&addr&"',color='"&color&"',indate='"&indate&"' where id=" & request("id")
response.redirect "viewdocuments.asp"
end if 
%>
<%
if trim(color)="" then
color=" "
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
sql="select * from document where id=" & request("id")
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
%>
<table border="0" cellspacing="1" cellpadding="5" bgcolor="#000000" width="65%" align="center">
  <form action="EDITDOCUMENTS.ASP" method="post" name="editdocuments">
    <tr bgcolor="#FFCC33"> 
      <td height="22" width="17%">文档标题：</td>
      <td colspan="3"> &nbsp;&nbsp; <input maxlength=45 name="subject" size=40 class="form" value="<%=rs("subject")%>"> 
      </td>
    </tr>
    <tr bgcolor="#FFCC33"> 
      <td height="22" width="17%">文档地址：</td>
      <td colspan="3"> &nbsp;&nbsp; <input maxlength=45 name="addr" size=40 class="form" value="<%=rs("addr")%>"></td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td height="22" valign="top">color</td>
      <td width="27%">&nbsp;&nbsp; <input maxlength=6 name="color" size=10 class="form" value="<%=rs("color")%>"></td>
      <td width="19%">添加日期：</td>
      <td width="37%"><input maxlength=6 name="indate" size=10 class="form" value="<%=rs("indate")%>"></td>
    </tr>
    <tr bgcolor="#FFFFCC"> 
      <td height="45" colspan="4"> <div align="center"> 
          <input type="submit" name="Submit" value="确 定">
          &nbsp;&nbsp; 
          <input type="reset" name="Submit" value="取 消">
          &nbsp;&nbsp; 
          <input type="button" name="Submit" value="后 退" onClick="history.go( -1 );return true;">
          <input type="hidden" name="action" value="modi">
          <input type="hidden" name="id" value="<%=request("id")%>">
        </div></td>
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
