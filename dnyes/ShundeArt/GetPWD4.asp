<%response.buffer="True"%>
<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkURL.asp"-->
<%
dim rs,sql
dim username,passwd
username=checkStr(request("username"))
passwd=trim(request.form("passwd"))
passwd2=trim(request.form("passwd2"))
if username="" or passwd="" or passwd2="" then
	Show_Err("�벻Ҫ���ԷǷ�����!!<br><br><a href='javascript:history.back()'>����</a>")
	response.end
end if
if passwd<>passwd2 then
	Show_Err("���ε����벻һ��,�벻Ҫ���ԷǷ�����!!<br><br><a href='javascript:history.back()'>����</a>")
	response.end
end if
passwd=md5(passwd)
sessionusername=session("sessionusername")     '�õ�session
sessionanswer=session("sessionanswer")         '�õ�session
sessionfullname=session("sessionfullname")     '�õ�session
if sessionusername<>username then              '�ж�session
	Show_Err("�벻Ҫ���ԷǷ�����!!<br><br><a href='javascript:history.back()'>����</a>")
	response.end
end if

set rs=Server.CreateObject("Adodb.Recordset")
sql="select * from "& db_User_Table &" where  "& db_User_Name &"='"&username&"'"
rs.open sql,ConnUser,1,3
If rs.eof Then
	Show_Err("����û���û��ע���أ��뵽��ҳע��ɣ�<br><br><a href='javascript:history.back()'>����</a>")
	response.end
elseIf rs(db_User_Answer)<>sessionanswer Then         '�ж�session
		Show_Err("�벻Ҫ���ԷǷ�����!!<br><br><a href='javascript:history.back()'>����</a>")
		response.end
elseIf rs("fullname")<>sessionfullname Then     '�ж�session
			Show_Err("�벻Ҫ���ԷǷ�����!!<br><br><a href='javascript:history.back()'>����</a>")
			response.end
else
    rs(db_User_Password)=passwd
    rs.update
     session("sessionusername")=""    '����session
     session("sessionanswer")=""      '����session
     session("sessionfullname")=""    '����session
end if
rs.close
set rs=nothing
conn.close   
set conn=nothing   
 

%>
<html><head>
<title>ȡ������</title>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="style.css">
<script language="javascript">
<!--
if (parent.frames.length > 0) {
parent.location.href = location.href;
}
// --></script>
</head>
<!--#include file=User_Top.asp-->
<body topmargin="0" oncontextmenu="self.event.returnValue=false">
<div align="center"><br><br>
<center>
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="760" id="AutoNumber1">
<tr>
<td height="127" valign="top" width="100%"></center>
<div align="center">
<center>
<table border="0" cellpadding="5" width="100%" cellspacing="4" height="49">
<tr>
<td width="100%" height="1">
<p align="center">�������óɹ���</td>
</tr>
</table>
</center>
</div>
</td>
</tr>
<tr>
<td height="14" width="100%">
<p align="center"><a href="javascript:window.close()" class="1">�رմ���</a></td>
</tr>
</table>
</form>
</div>
</body>
</html>
<!--#include file=User_Bottom.asp-->