<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<% chkAdmin() %>
<% Response.buffer=true %>
<!--
	COCOON Counter 6
	Develop by Cocoon Studio. [ www.ccopus.com ]
	Product by Sunrise_Chen. (MSN:sunrise_chen@msn.com)
	
	Please don't remove this information, thanks.
-->
<html>
<head>
<title>管理工具 - COCOON Counter 6</title>
<style>
body,form,legend,td,fieldset {font-size:9pt;cursor:default}
input {border-width:1px;font-size:9pt;}
</style>
<script language=javascript src="../script/nomenu.js"></script>
</head>
<body style="background-color:buttonface;padding:0;margin:5;border:0">
<form method="post" style="margin:0">
<fieldset style="width:100%;height=75px">
<legend>删除用户</legend>
<div style="width:100%;height:75px;overflow:auto">
<table width="100%" height="100%"><tr><td align=left valign=top>
<%
	AdminPwd = "sunrise"
	var sUID = (Request.QueryString("uid") + "").replace(/\'/g,"''")
	if(sUID.length=0) Resposne.End()

	oConn.Open()
	var oRs = oConn.execute("select SiteId from t_Site where uid='" + sUID + "'")
	if(oRs.EOF){
		oConn.Close()
		Response.write("用户不存在")
		Response.end()
	}
	var sId = oRs.fields.item(0).value
	oConn.execute("delete from t_Visit where SiteId=" + sId)
	Response.write("删除最后访问信息 ... 完成.<br>\n")
	Response.flush()
	oConn.execute("delete from t_Region where SiteId=" + sId)
	Response.write("删除地区统计信息 ... 完成.<br>\n")
	Response.flush()
	oConn.execute("delete from t_Page where SiteId=" + sId)
	Response.write("删除页面统计信息 ... 完成.<br>\n")
	Response.flush()
	oConn.execute("delete from t_Link where SiteId=" + sId)
	Response.write("删除上网方式统计信息 ... 完成.<br>\n")
	Response.flush()
	oConn.execute("delete from t_RefPage where SiteId=" + sId)
	Response.write("删除引用页统计信息 ... 完成.<br>\n")
	Response.flush()
	oConn.execute("delete from t_Keyword where SiteId=" + sId)
	Response.write("删除关键字统计信息 ... 完成.<br>\n")
	Response.flush()
	oConn.execute("delete from t_Client where SiteId=" + sId)
	Response.write("删除客户端统计信息 ... 完成.<br>\n")
	Response.flush()
	oConn.execute("delete from t_Count where SiteId=" + sId)
	Response.write("删除时间统计信息 ... 完成.<br>\n")
	Response.flush()
	oConn.execute("delete from t_Site where SiteId=" + sId)
	Response.write("删除基本信息 ... 完成.<br>\n")
	Response.flush()
	oConn.Close()
/*
	Response.write("压缩数据库 ... ")
	try{
		Server.execute("Compress.asp")
		Response.write("完成.<br>\n")
	}catch(e){
		Response.write("失败: <font color=red>"+e.description+"</font>.<br>\n")
	}
*/
%>
</td></tr></table>
</div>
</fieldset>
</form>
</body>
</html>