<%
'要找回密码,除要一步一步的判断用户的信息回答是否正确外,
'关键还要判断用户的每一步是否完成,是否存在跳过的步骤,和外部提交,
'在此修改中运用了seesion判断用户每一步是否完成,
'最后判断session的存在和正确性,再消除session,
'**TUPUNCO--2003/12/9修改**
%>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkURL.asp"-->
<%
dim rs,sql
dim username
username=checkstr(request.form("username"))
if username="" then
  %>
  <script language="javascript">
   alert("请不要尝试非法操作!!")
   location.href="javascript:history.back()"
  </script>
<%
Response.End
end if

set rs=Server.CreateObject("Adodb.Recordset")
sql="select  "& db_User_Name &","& db_User_Question &" from "& db_User_Table &" where  "& db_User_Name &"='"&username&"'"
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
session("sessionusername")=username  'session
%>
<html>

<head>
<!--#include file=User_Top.asp-->
<title>取回密码 | 回答问题</title>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="style.css">
<script language="javascript">
<!--
if (parent.frames.length > 0) {
parent.location.href = location.href;
}
function form1_onsubmit() {
if (document.form1.name.value=="")
	{
	  alert("请输入您的用户名。")
	  document.form1.name.focus()
	  return false
	 }
}
// --></script>
</head>

<body topmargin="0" oncontextmenu="self.event.returnValue=false">

<div align="center"><br><br>
<center>
<form method="POST" name="form1" language="javascript" onsubmit="return form1_onsubmit()" action="getpwd3.asp?username=<%=username%>">
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="760" id="AutoNumber1">
<tr>
<td width="760"></center>
<div align="center">
<center>
<table border="0" cellpadding="0" width="99%" cellspacing="0" height="1" style="border-collapse: collapse" bordercolor="#111111"><%if rs(db_User_Question)<>"" then%>
<tr>
<td width="760">
<p align="center">第二步：请回答下列问题。</td>
</tr>
<tr>
<td width="760">
<p align="center">问题：<font color="red"><%=rs(db_User_Question)%></font>
<%     
conn.close     
set conn=nothing     
%></td> 
</tr> 

<tr> 
<td width="760">
<p align="center">答案：<input type="text" name="answer" size="32" style="font-family: 宋体; font-size: 9pt;></td>
</tr>
<tr>
<td width="760">
<p align="center">您的真实姓名：<input type="text" name="fullname" size="23" style="font-family: 宋体; font-size: 9pt;></td>
</tr>
<tr>
<td width="760">
<p align="center"><input type="submit" value=" 到下一步 " name="submit" style="font-family: 宋体; font-size: 9pt;></td>
</tr><%else%>
<tr>
<td width="760"align=center>您在个人资料中设置的提示问题为空，所以不能找回密码！如果您真的忘记了密码，请与<a href="mailto:<%=email%>">管理员</a>联系！</td>
</tr>
<%end if%>
<tr>
<td width="760"></td>
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