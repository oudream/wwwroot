<!--#include file="conn.asp" -->
<!--#include file="ConnUser.asp" -->
<!--#include file="config.asp" -->
<!--#include file="char.inc"-->
<!--#include file="ChkUser.asp"-->

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
<%
response.cookies(Forcast_SN)("ManageUserName")=""
response.cookies(Forcast_SN)("ManageKEY")=""
response.cookies(Forcast_SN)("ManagePurview")=""
response.cookies(Forcast_SN)("ManageFullName")=""
response.cookies(Forcast_SN)("ManageReglevel")=""
showerr=request("showmess")

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
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%><%=ver%></title>
</head>
<body bgcolor="#c1d2eb" topmargin="0" marginheight="0">
<br>
<p>&nbsp;</p>
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<form method="POST" action="ChkManageLogin.asp" name="UserLogin" onSubmit="return CheckFormUserLogin();">
<tr >
	<td height="67" colspan="2" align="center" background="IMAGES/top1.gif" >
		<b><font color="#ffffff" ><%=copyright%>&nbsp;<p><%=version%>&nbsp;<%=ver%></font></b>
	</td>
</tr>
<tr>
	<td height="37" colspan="2" align="center" class="TDtop">
		<span ><font color=red>┊ <strong>管 理  登 录</strong> ┊</font></span>
	</td>
</tr>
<tr>
	<td width="45%" align="right">用户名：</td>
	<td width="55%" align="left">
		<input name="UserName" size="15" font face="宋体" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
<tr>
	<td align="right">密&nbsp;&nbsp;码：</td>
	<td align="left">
		<input type="password" name="Passwd" size="15" font face="宋体" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
<tr>
	<td align="right">验证码：</td>
	<td align="left">
		<input type="text" name="verifycode" size="15" font face="宋体" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<b><span><font color=#000000><%=getcode1()%></font></span></b>
	</td>
</tr>
<%if showerr<>"" then%>
<tr>
	<td colspan="2" align="center">
		<p><font size=3 color=red><%=showerr%></font></p>
	</td>
</tr>
<%end if%>
<tr>
	<td colspan="2" align="center">
		<p>
			<input type="submit" name="Submit" value="确定" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="确定">&nbsp;
			<input type="reset" name="Submit2" value="重置" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title="重输">
		</p>
	</td>
</tr>
<tr >
	<td colspan="2" align="center" background="IMAGES/top2.gif">&nbsp;</td>
</tr>
</form>
</table>
</body>
</html>
<%response.end%>