<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->

<html>

<head>
<!--#include file=User_Top.asp-->
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title> ȡ������ | ����</title><link rel="stylesheet" type="text/css" href="style.css">
</head>

<body topmargin="0" oncontextmenu="self.event.returnValue=false">
<div align="center">
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="760" id="AutoNumber1">
<tr class="TDtop">
	<td colspan=2 width="100%" height="25" align="center">
	<br><br><br><br>
	<p align="center">����<%=request("content")%>��д������<a href="javascript:history.back()">����</a>������д��
	<br><br><br><br>
	</td>
</tr>
</table>
</div>
</body>

</html>
<!--#include file=User_Bottom.asp-->