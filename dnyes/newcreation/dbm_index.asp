<%
if request("err")="uid_err" then
%>
<SCRIPT LANGUAGE=JavaScript>
alert(' 名称不对 . ');
window.close();
</SCRIPT>
<%
end if
%>
<%
if request("err")="pwd_err" then
%>
<SCRIPT LANGUAGE=JavaScript>
alert(' 密码不对 . ');
window.close();
</SCRIPT>
<%
end if
%>
<head>
<title>Control panel - design by dnyes.com</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="dbm_css.css" type=text/css rel=stylesheet>
</head>

<frameset cols="230,*" frameborder="NO" border="0" framespacing="0" rows="*"> 
  <frame name="leftFrame" scrolling="yes" noresize src="dbm_left.asp">
  <frame name="mainFrame" src="dbm_adminlogin.asp">
</frameset>
<noframes><body bgcolor="#FFFFFF">

</body></noframes>
</html>