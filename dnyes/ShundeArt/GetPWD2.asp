<%
'Ҫ�һ�����,��Ҫһ��һ�����ж��û�����Ϣ�ش��Ƿ���ȷ��,
'�ؼ���Ҫ�ж��û���ÿһ���Ƿ����,�Ƿ���������Ĳ���,���ⲿ�ύ,
'�ڴ��޸���������seesion�ж��û�ÿһ���Ƿ����,
'����ж�session�Ĵ��ں���ȷ��,������session,
'**TUPUNCO--2003/12/9�޸�**
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
   alert("�벻Ҫ���ԷǷ�����!!")
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
alert("����û���û��ע���أ��뵽��ҳע��ɣ�")
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
<title>ȡ������ | �ش�����</title>
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
	  alert("�����������û�����")
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
<p align="center">�ڶ�������ش��������⡣</td>
</tr>
<tr>
<td width="760">
<p align="center">���⣺<font color="red"><%=rs(db_User_Question)%></font>
<%     
conn.close     
set conn=nothing     
%></td> 
</tr> 

<tr> 
<td width="760">
<p align="center">�𰸣�<input type="text" name="answer" size="32" style="font-family: ����; font-size: 9pt;></td>
</tr>
<tr>
<td width="760">
<p align="center">������ʵ������<input type="text" name="fullname" size="23" style="font-family: ����; font-size: 9pt;></td>
</tr>
<tr>
<td width="760">
<p align="center"><input type="submit" value=" ����һ�� " name="submit" style="font-family: ����; font-size: 9pt;></td>
</tr><%else%>
<tr>
<td width="760"align=center>���ڸ������������õ���ʾ����Ϊ�գ����Բ����һ����룡�����������������룬����<a href="mailto:<%=email%>">����Ա</a>��ϵ��</td>
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
<p align="center"><a href="javascript:window.close()" class="1">�رմ���</a></td>
</tr>
</table>
</form>
</div>

</body>

</html>
<!--#include file=User_Bottom.asp-->