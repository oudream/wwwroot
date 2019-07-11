<%response.buffer="True"%>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkURL.asp"-->
<%
dim rs,sql
dim username,fullname,answer
answer=md5(trim(request.form("answer")))
fullname=checkstr(request.form("fullname"))
username=checkstr(request("username"))
if username="" or fullname="" or answer="" then
  %>
  <script language="javascript">
   alert("请不要尝试非法操作!!")
   location.href="javascript:history.back()"
  </script>
<%
Response.End
end if
set rs=Server.CreateObject("Adodb.Recordset")
sql="select * from "& db_User_Table &" where  "& db_User_Name &"='"&username&"'"
rs.open sql,ConnUser,1,1
If rs.eof Then
%>
<script language="javascript">
alert("这个用户还没有注册呢，请到首页注册吧！")
location.href="javascript:history.back()"
</script>
<%
Response.End
End If

If rs(db_User_Answer)<>answer Then
Response.redirect"getpwderr.asp?content=答案"
Response.End
End If

If rs("fullname")<>fullname Then
Response.redirect"getpwderr.asp?content=真实姓名"
Response.End
End If

session("sessionanswer")=answer              'session
session("sessionfullname")=fullname         'session
%>
<html>

<head>
<!--#include file=User_Top.asp-->
<title>取回密码</title>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="style.css">
<script language="javascript">
<!--
if (parent.frames.length > 0) {
parent.location.href = location.href;
}
// --></script>
<script LANGUAGE="javascript">
<!--
function FrmAddLink_onsubmit() {
 if (document.FrmAddLink.passwd.value=="")
	{
	  alert("对不起，请您输入密码！")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd.value.length < 4)
	{
	  alert("为了安全，您的密码应该长一点！")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd.value.length > 16)
	{
	  alert("您的密码太长了吧！")
	  document.FrmAddLink.passwd.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd2.value=="")
	{
	  alert("对不起，请您输入验证密码！")
	  document.FrmAddLink.passwd2.focus()
	  return false
	 }
else if (document.FrmAddLink.passwd2.value !== document.FrmAddLink.passwd.value)
	{
	  alert("对不起，您两次输入的密码不一致！")
	  document.FrmAddLink.passwd2.focus()
	  return false
	 }
}

//-->
</script>
</head>

<body topmargin="0" oncontextmenu="self.event.returnValue=false">

<div align="center"><br><br>
<center>
<form method="POST" name="FrmAddLink" language="javascript" onsubmit="return FrmAddLink_onsubmit()" action="getpwd4.asp?username=<%=username%>">
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="760" id="AutoNumber1">
<tr>
<td width="760"></center>
<div align="center">
<center>
              <table border="0" cellpadding="5" width="99%" cellspacing="4" height="49">
                <tr>
<td width="760">
<p align="center">第三步：重新设置密码。</td>
</tr>
<tr> 
<td width="760">
<p align="center">新&nbsp;密&nbsp;码：<input type="password" name="passwd" size="32" maxlenght=16 style="font-family: 宋体; font-size: 9pt;></td>
</tr>
<tr> 
<td width="760">
<p align="center">验证密码：<input type="password" name="passwd2" size="32" maxlenght=16 style="font-family: 宋体; font-size: 9pt;></td>
</tr>
<tr>
<td width="760">
<p align="center"><input type="submit" value=" 设置密码 " name="submit" style="font-family: 宋体; font-size: 9pt;></td>
</tr>
</table>
</center>
</div>
</td>
</tr>
<tr>
<td width="760">
<p align="center"><a href="javascript:window.close()" class="1">关闭窗口</a></td>
</tr>
</table>
</form>
</div>

</body>

</html>
<!--#include file=User_Bottom.asp-->
<%   
conn.close   
set conn=nothing   
 
%>