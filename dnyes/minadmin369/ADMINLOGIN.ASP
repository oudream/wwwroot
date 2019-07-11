<!--#include file="../mdf.asp" -->
<%
if request("logout")<>"" then 
session("y_eshop_adminuser")=""
session("y_eshop_aleave")=""
response.redirect "adminlogin.asp"
end if
%>
<%

if request("action")="login" then
auid=request("auid")
apwd=request("apwd")
apwd=md5(apwd)

Set Conn=Server.CreateObject("ADODB.Connection") 
Connstr="DBQ="+server.mappath("../mymefaq5411jkjkh/asdewr2482f/safkk355f.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)}"
Conn.Open connstr 
sql="select * from adminuser where auid='"&auid&"'"
set rs=server.createobject("adodb.recordset") 
rs.open sql,conn,1,1
if rs.eof or rs.bof then
response.redirect "adminlogin.asp"
else
if apwd<>rs("apwd") then 
response.redirect "adminlogin.asp"
else
session("y_eshop_adminuser")=auid
session("y_eshop_aleave")=rs("aleave")
response.redirect "adminlogin.asp"
rs.close
set rs=nothing
conn.close
set conn=nothing
end if
end if
end if
%>


<html>
<head>
<title>系统管理中心</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../inc/southdns.css" type="text/css">
</head>
<body bgcolor="#FFFFCC" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="1" cellpadding="2" bgcolor="#000000" vspace="0" hspace="0">
  <tr bgcolor="#FFCC33"> 
    <td height="27">.:: 系统管理中心</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="70%">
  <tr>
    <td align="center" valign="middle"><b><font size="3" color="#000066">
      <br>
      <br>
      <br>
      <br>
      欢迎进入系统管理中心</font></b><br>
      <br>
      <br>
      <br>
      <br>
<%if session("y_eshop_adminuser")="" then%>
      <table border="0" cellspacing="1" cellpadding="4" bgcolor="#000000">
        <tr bgcolor="#FFCC33"> 
          <td colspan="2">Admin登陆</td>
        </tr>
        <form name="Login" method="post" action="ADMINLOGIN.ASP">
          <tr bgcolor="#FFFF99"> 
            <td align="right">用户名:</td>
            <td> 
              <input type="text" name="auid" class="form">
            </td>
          </tr>
          <tr bgcolor="#FFFF99"> 
            <td align="right">密码:</td>
            <td> 
              <input type="password" name="apwd" class="form">
            </td>
          </tr>
          <tr bgcolor="#FFFF99"> 
            <td colspan="2" align="center"> 
              <input type="submit" name="Submit" value="确 定">
          <input type="hidden" name="action" value="login"></td>
          </tr></form>
      </table>
<%
else
response.write "恭喜，恭喜！<br><br>您已经成功登陆！<br><br><br><br>"
end if
%>


</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td height="50">&nbsp;</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="center"></td>
  </tr>
</table>
</body>
</html>
