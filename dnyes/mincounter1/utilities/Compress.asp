<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<% chkAdmin() %>
<%
	//Response.Expires = -1
	
	var oFso = new ActiveXObject("Scripting.FileSystemObject");
	var sDbPath = Server.MapPath(DbPath);
	var oFile = oFso.getFile(sDbPath);
	var iBeginSize = oFile.size;
	var iEndSize = 0;
	
	if(Request("Compress").count>0){
		var oJro = new ActiveXObject("jro.JetEngine");	
		var sTmpDbPath = Server.MapPath(DbPath + ".bak");
		var sTmpConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+sTmpDbPath+";Jet OLEDB:Database Password="+DbPassword+";";
		if(oFso.FileExists(sTmpDbPath)) oFso.DeleteFile(sTmpDbPath);
		oJro.CompactDatabase(oConn.ConnectionString,sTmpConnString);
		oFso.CopyFile(sTmpDbPath,sDbPath);
		oFso.DeleteFile(sTmpDbPath);
		oFile = oFso.getFile(sDbPath);
		iEndSize = oFile.size;
		oJro = null;
		oFso = null;
	}

%>
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
<fieldset style="width:100%;height:75px">
<legend>数据库压缩</legend>
<table width="100%" height="100%"><tr><td align=left>
	　　当前数据库大小: <%=Math.round((iBeginSize/1024/1024)*100)/100%> MB.
	
	<% if(Request.form("Compress").count>0){ %>
	<br>　　压缩后的大小为: <%=Math.round((iEndSize/1024/1024)*100)/100%> MB.<br>
	<div align=right>
		<img src="../images/blank.gif" height=22 width=1>
		<input type=button value="　关闭　" onclick="window.close()">
		<input type=button value="　返回　" onclick="location.href='../admin/default.asp'">
	</div>
	<% }else{ %><br><font color=gray>　　Tips: 请尽量在流量较小的时候进行此项操作。</font><br>
	<div align=right>
		<img src="../images/blank.gif" height=22 width=1>
		<input type=submit name=Compress value="开始压缩" style="font-size:9pt">
		<input type=button value="　返回　" onclick="location.href='../admin/default.asp'">
	</div>
	<% } %>
</td></tr></table>
</fieldset>
</form>
</body>
</html>