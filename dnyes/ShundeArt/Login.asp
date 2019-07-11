<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<%
response.cookies(Forcast_SN)("UserName")=""
response.cookies(Forcast_SN)("KEY")=""
response.cookies(Forcast_SN)("purview")=""
response.cookies(Forcast_SN)("fullname")=""
response.cookies(Forcast_SN)("reglevel")=""
response.cookies(Forcast_SN)("ViewUrl")="Admin_login.asp"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title>后台入口</title>
<script language="JavaScript">
<!--

if (self != top) top.location.href = window.location.href

//-->
</script>

<script language=javascript>
function CheckFormUserLogin()
{
	if(document.UserLogin.UserName.value=="")
	{
		alert("请输入用户名！");
		document.UserLogin.UserName.focus();
		return false;
	}
	if(document.UserLogin.Passwd.value == "")
	{
		alert("请输入密码！");
		document.UserLogin.Passwd.focus();
		return false;
	}
	if(document.UserLogin.verifycode.value == "")
	{
		alert("请输入验证码！");
		document.UserLogin.verifycode.focus();
		return false;
	}
}
</script>

<style type="text/css">
<!--
.style1 {
	font-size: 10.5pt;
	font-weight: bold;
}
-->
</style>
</head>
<body bgcolor="#CCCCCC" background="images/linebg1.gif" topmargin="0" marginheight="0">
<br>
<p>&nbsp;</p>
<form method="POST" action="ChkLogin.asp" name="UserLogin" onSubmit="return CheckFormUserLogin();">
  <table width="750" border="0" align=center cellpadding="0" cellspacing="2" bgcolor="#CCCCCC">
    <tr> 
      <td height="67" align="center" bgcolor="#FFFFFF"><strong><font size="3">后 
        台 入 口</font></strong></td>
    </tr>
    <tr> 
      <td height="37" align="center" bgcolor="#FFFFFF"><span class="style1"><font color=#FFFFFF>用 
        户 登 录</font></span></td>
    </tr>
    <tr> 
      <td height="176" bgcolor="#FFFFFF"> <table width="350" border="0" align="center" cellpadding="6" cellspacing="2" bgcolor="#CCCCCC">
          <tr> 
            <td bgcolor="#FFFFFF">用户名： </td>
            <td bgcolor="#FFFFFF"><input name="UserName" size="15" font face="宋体" style="font-size: 9pt; background-color:#EAEAF4"></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF">密&nbsp;&nbsp;码： </td>
            <td bgcolor="#FFFFFF"><input type="password" name="Passwd" size="15" font face="宋体" style="font-size: 9pt; background-color:#EAEAF4"></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF">验证码： </td>
            <td bgcolor="#FFFFFF">
              <%
		Function getcode1()
			Dim test
			On Error Resume Next
			Set test=Server.CreateObject("Adodb.Stream")
			Set test=Nothing
			If Err Then
				Dim zNum
				Randomize timer
				zNum = cint(8999*Rnd+1000)
				Session("verifycode") = zNum
				getcode1= Session("verifycode")		
			Else
				getcode1= "<img src=""getcode.asp"">"		
			End If
		End Function
		%>
              <input type="text" name="verifycode" size="15" font face="宋体" style="font-size: 9pt; background-color:#EAEAF4"> 
              <b><span><font color=#000000><%=getcode1()%></font></span></b></td>
          </tr>
          <tr> 
            <td colspan="2" bgcolor="#FFFFFF"> <p> 
                <input type="submit" name="Submit" value="确定" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="确定">
                &nbsp; 
                <input type="reset" name="Submit2" value="重输" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="重输">
              </p></td>
          </tr>
        </table></td>
    </tr>
  </table>
</form>
</body>
</html>