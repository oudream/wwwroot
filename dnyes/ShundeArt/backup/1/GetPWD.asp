<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html>
<head>
<!--#include file=User_Top.asp-->
<title>ȡ������</title>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="style.css">
<script language="javascript">
<!--
if (parent.frames.length > 0) {
parent.location.href = location.href;
}
function form1_onsubmit() {
if (document.form1.username.value=="")
	{
	  alert("�û�����û�У�����ô���һ����������أ�")
	  document.form1.username.focus()
	  return false
	 }
}
// --></script>
</head>
<body topmargin="0" oncontextmenu="self.event.returnValue=false">
<div align="center"><br><br>
<form method="POST" name="form1" language="javascript" onsubmit="return form1_onsubmit()" action="getpwd2.asp">
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="760" id="AutoNumber1">
<center>
</center><tr><td valign="top">
<center>
<p></p>
</center>
<div align="center">
<center>
<table border="0" cellpadding="0" width="760" cellspacing="0">
<tr>
<td width="760">
<p align="center">���������ˣ��ش𼸸������Ϳ�������������������</td>
</tr>
<tr>
<td width="760">
<p align="center">��һ������д�����û�����</td>
</tr>
<tr>
<td width="760">
<p align="center"><input type="text" name="username" size="30" class="form" style="font-family: ����; font-size: 9pt;></td>
</tr>
<tr>
<td width="760">
<p align="center"><input type="submit" value=" ����һ�� " name="submit" style="font-family: ����; font-size: 9pt;></td>
</tr>
<tr>
<td width="100%" height="188"></td>
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