<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Config.asp"-->
<%
Dim Msg,Uname,Upass
Uname = replace(trim(Request.Form("Uname")),"'","")
Upass = replace(trim(Request.Form("Upass")),"'","")

	If Request.QueryString("Action") = "Check" Then Check Uname,Upass
	Sub Check(Uname,Upass)
	Set rs = myConn.execute("Select Adminname,AdminPass From Config")
	If rs.Eof Then
	Msg = "ϵͳ���ݿ��������ϸ������ݿ�ṹ�����ݼ�¼"
	Exit Sub
	End If
	If Uname <> rs(0) or Upass <> rs(1) Then
	Msg = "�ʻ����������"
	Exit Sub
	End If
	Session("Admin") = "OK"
	Msg = "�����֤ͨ��"
	End Sub

If msg <> "" then
If Session("Admin") = "OK" Then
Response.Write("<meta http-equiv=refresh content='3;URL=Admin_List.asp'>"&Msg&"<br>��ҳ����3���ڷ���<BR>�����������û�з�Ӧ����<a href=Admin_List.asp>����˴�����</a>")
Response.End()
Else
Response.Write("<meta http-equiv=refresh content='3;URL=Admin_Login.asp'>"&Msg&"<br>��ҳ����3���ڷ���<BR>�����������û�з�Ӧ����<a href=Admin_Login.asp>����˴�����</a>")
Response.End()
End If
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>:::������ϴ����� ::: ��ʾ�汾</title>
<link href="upstyle.css" rel="stylesheet" type="text/css">
<script src="Images/up.Js"></script>
</head>

<body background="../images/mainbg.gif" leftmargin="0" topmargin="0">
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#003366">
  <tr> 
    <td height="75">&nbsp;</td>
  </tr>
  <tr> 
    <td height="22" bgcolor="#FF9900" class="text">��ǰλ�ã�:::������ϴ�����Ver1.2 ::: ��ʾ�汾 
      &gt;&gt;&gt; ��̨����&nbsp;��&nbsp;<a href="./">�����ϴ���ҳ</a></td>
  </tr>
  <tr>
    <td height="2" bgcolor="#cccccc"></td>
  </tr>
</table>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="778" height="200">
        <param name="movie" value="FLASH_AD.swf">
        <param name="quality" value="high">
        <embed src="FLASH_AD.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="778" height="200"></embed></object></td>
  </tr>
  <tr>
    <td height="1" bgcolor="#CCCCCC"></td>
  </tr>
</table>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#0153A3">
<form name="Check" action="Admin_Login.asp?Action=Check" method="post">
  <tr> 
    <td height="25">
      <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
          <tr align="center" class="text"> 
            <td height="1" colspan="2" bgcolor="#0153A3">��̨����-��½ϵͳ</td>
          </tr>
          <tr class="text"> 
            <td width="45%"> 
              <div align="right">����Ա�ʻ���</div></td>
            <td><input name="uname" type="text" class="TextBoxT" id="uname" size="15">
            </td>
          </tr>
          <tr class="text"> 
            <td>
<div align="right">����Ա���룺</div></td>
            <td><input name="upass" type="password" class="TextBoxT" id="upass" size="15">
            </td>
          </tr>
          <tr class="text"> 
            <td height="1" align="right" bgcolor="#0153A3">&nbsp;</td>
            <td height="1" bgcolor="#0153A3"> <input type="submit" name="Submit" value="��½"></td>
          </tr>
        </table>
	  <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
        <tr class="text"> 
            <td align="right" class="heading">&nbsp;&nbsp;&nbsp;&nbsp;</td>
        </tr>
      </table></td>
  </tr></form>
</table>

<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#0153A3">
  <tr> 
    <td height="25">&nbsp;</td>
  </tr>
</table>
</body>
</html>