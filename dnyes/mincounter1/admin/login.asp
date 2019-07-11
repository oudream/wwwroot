<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<%
	if(isPostBack()){
		doLoginAdmin(Request.form("adminpwd")+"")
		if(chkAdmin()){
			doAlert("","../admin/default.asp")
		}
	}
%>
<html>
<head>
<title>管理工具 - COCOON Counter 6</title>
<style>
body,input,form,legend,td {font-size:9pt;}
</style>
</head>
<body style="background-color:buttonface;padding:0;margin:5;border:0">
<form method="post" style="margin:0">
<fieldset style="width:100%;height=75px">
<legend>管理员登录</legend>
<table width="100%" height="100%"><tr><td align=center>
	管理员密码：
	<input type="password" name="adminpwd" onkeydown="event.cancelBubble=true">
	<input type="submit" name="submit" value="登录">
</td></tr></table>
</fieldset>
</form>
</body>
</html>