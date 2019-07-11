<!--#include file="SHEI45FECSA.ASP" -->
<%on error resume next%>
<%
'修改公告
if request("action")="modi" then

color=request("color")
subject=request("subject")
newstype=request("newstype")
message=replace(request("message"),">","&gt;")
message=replace(message,"<","&lt;")
message=replace(message,"'","’")
conn.execute "update message set subject='"&subject&"',message='"&message&"',newstype='"&newstype&"',color='"&color&"' where id=" & request("id")
response.redirect "viewnews.asp"
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
sql="select * from message where id=" & request("id")
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
%> 
<table border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" width="65%" align="center">
  <form action="EDITNEWS.ASP" method="post" name="news">
    <tr bgcolor="#FFCC33"> 
      <td height="22" width="17%">&nbsp;公告标题：</td>
      <td colspan="3"> &nbsp; <input maxlength=45 name=subject size=40 class="form" value="<%=rs("subject")%>"> 
      </td>
    </tr>
    <tr bgcolor="#FFCC33"> 
      <td height="22" width="17%">公告类型：</td>
      <td width="30%"> &nbsp;&nbsp;
        <select name="newstype">
          <option value="<%=rs("newstype")%>"><%=rs("newstype")%></option>
          <option value="行业信息">行业信息</option>
          <option value="本站动态">本站动态</option>
          <option value="客服中心">客服中心</option>
        </select> </td>
      <td width="14%">颜色：</td>
      <td width="39%"><input maxlength=6 name=color id="color" size=20 class="form" value="<%=rs("color")%>"></td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td height="22" valign="top"><br> &nbsp;广告内容：</td>
      <td colspan="3"><br> &nbsp; <textarea name="message" cols="50" class="form" rows="8">
<%=rs("message")%>
</textarea> 
        <br> <br> </td>
    </tr>
    <tr bgcolor="#FFFF99"> 
      <td height="45" colspan="4"> <div align="center"> 
          <input type="submit" name="Submit" value="确 定">&nbsp;&nbsp;
          <input type="reset" name="Submit" value="取 消">&nbsp;&nbsp;
          <input type="button" name="Submit" value="后 退" onClick="history.go( -1 );return true;">
          <input type="hidden" name="action" value="modi">
          <input type="hidden" name="ID" value="<%=request("id")%>">
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
