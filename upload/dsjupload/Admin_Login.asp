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
	Msg = "系统数据库错误，请仔细检查数据库结构及数据记录"
	Exit Sub
	End If
	If Uname <> rs(0) or Upass <> rs(1) Then
	Msg = "帐户或密码错误"
	Exit Sub
	End If
	Session("Admin") = "OK"
	Msg = "身份验证通过"
	End Sub

If msg <> "" then
If Session("Admin") = "OK" Then
Response.Write("<meta http-equiv=refresh content='3;URL=Admin_List.asp'>"&Msg&"<br>本页将在3秒内返回<BR>如果你的浏览器没有反应，请<a href=Admin_List.asp>点击此处返回</a>")
Response.End()
Else
Response.Write("<meta http-equiv=refresh content='3;URL=Admin_Login.asp'>"&Msg&"<br>本页将在3秒内返回<BR>如果你的浏览器没有反应，请<a href=Admin_Login.asp>点击此处返回</a>")
Response.End()
End If
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>:::丁书记上传程序 ::: 演示版本</title>
<link href="upstyle.css" rel="stylesheet" type="text/css">
<script src="Images/up.Js"></script>
</head>

<body background="../images/mainbg.gif" leftmargin="0" topmargin="0">
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#003366">
  <tr> 
    <td height="75">&nbsp;</td>
  </tr>
  <tr> 
    <td height="22" bgcolor="#FF9900" class="text">当前位置：:::丁书记上传程序Ver1.2 ::: 演示版本 
      &gt;&gt;&gt; 后台管理&nbsp;—&nbsp;<a href="./">返回上传首页</a></td>
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
            <td height="1" colspan="2" bgcolor="#0153A3">后台管理-登陆系统</td>
          </tr>
          <tr class="text"> 
            <td width="45%"> 
              <div align="right">管理员帐户：</div></td>
            <td><input name="uname" type="text" class="TextBoxT" id="uname" size="15">
            </td>
          </tr>
          <tr class="text"> 
            <td>
<div align="right">管理员密码：</div></td>
            <td><input name="upass" type="password" class="TextBoxT" id="upass" size="15">
            </td>
          </tr>
          <tr class="text"> 
            <td height="1" align="right" bgcolor="#0153A3">&nbsp;</td>
            <td height="1" bgcolor="#0153A3"> <input type="submit" name="Submit" value="登陆"></td>
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