<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<% chkAdmin() %>
<html>
<head>
<title>管理工具 - COCOON Counter 6</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<style>
body,form,legend,td,fieldset {font-size:9pt;cursor:default}
input {border-width:1px;font-size:9pt;}
</style>
<script language=javascript src="../script/nomenu.js"></script>
</head>
<body style="background-color:buttonface;padding:0;margin:5;border:0;overflow:auto" scroll=auto>
<form method="post" style="margin:0;width:100%;height:84%" action="proc.asp">
<fieldset style="width:100%;height:100%">
<legend>用户管理工具</legend>
<div style="width:100%;height:90%;padding-left:10px;overflow:auto;">
<div align=right style="color:red">* 该操作是不可逆的，请慎重！</div>
<hr size="1" width="98%">
<span style='width:19%;text-align:left'>&nbsp;用户名</span>
<span style='width:19%;text-align:left'>&nbsp;密码</span>
<span style='width:38%;text-align:left'>&nbsp;网址</span>
<span style='width:18%;text-align:right'>操作</span>
<br>
<%
	oConn.open()
	var oRs = oConn.execute( "select uid,pwd,siteurl from t_Site order by siteid" );
	while(!oRs.EOF){
		Response.write("<span style='width:19%'>&nbsp;"+oRs(0).value+"</span>\n");
		Response.write("<span style='width:19%'>&nbsp;"+oRs(1).value+"</span>\n");
		Response.write("<span style='width:38%'>&nbsp;"+oRs(2).value+"</span>\n");
		Response.write("<span style='width:18%;text-align:right'><a href='../utilities/deleteUser.asp?uid="+oRs(0).value+"' onclick='return confirm(\"确定要删除吗？\")'>删除</a></span>");
		Response.write("<br>\n");
		oRs.moveNext();
	}
	oConn.close();
%>
</td></tr></table>
</div>
</fieldset>
<!--#include file="toolbar.asp"-->
</form>
</body>
</html>