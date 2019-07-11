<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<%
	iDeleteCount = 5;

	if(Request.ServerVariables("REQUEST_METHOD")!="POST")
		Response.end();
	
	chkAdmin();
%>
<html>
<head>
<title>管理工具 - COCOON Counter 6</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<style>
body,form,legend,td,fieldset {font-size:9pt;cursor:default}
input {border-width:1px;font-size:9pt;}
label {cursor:hand;}
.content{padding-left:20px;}
</style>
<script language=javascript src="../script/nomenu.js"></script>
</head>
<body style="background-color:buttonface;padding:0;margin:5;border:0;overflow:auto" scroll=auto>
<form method="post" style="margin:0;width:100%;height:84%">
<fieldset style="width:100%;height:100%">
<legend>管理工具</legend>
<div style="width:100%;height:90%;padding-left:10px;overflow:auto;">
<%
	var sOldPwd = Request.form("OldPwd")+'';
	var sNewPwd = Request.form("NewPwd")+'';
	var sCfmPwd = Request.form("CfmPwd")+'';
	Response.write("<br><b>修改密码</b> ... ");
	if(!sOldPwd){
		Response.write("未操作.<br>");
	}else{
		var s = read("../_inc/common.asp");
		//sOldPwd = sOldPwd.replace(/\\/,"\\\\").replace(/\"/g,'\\\"');
		if(sOldPwd!=AdminPwd){Response.write("操作失败，原因：原始密码不正确。<br>")}
		else{
			if(sNewPwd!=sCfmPwd){Response.write("操作失败，原因：两次输入的密码不相同。<br>")}
			else{
				var re = new RegExp('AdminPwd.+?=.+?".+?"[ \t]*$','gm');
				sNewPwd = sNewPwd.replace(/\\/,"\\\\").replace(/\"/g,'\\\"');
				s=s.replace(re,'AdminPwd			=	"'+sNewPwd+'"');
				if(write("../_inc/common.asp",s)) Response.write("操作失败，原因：你的服务器禁止了FSO组件。<br>");
				else Response.write("完成。<br>")
			}
		}
	}

	Response.write("<br><b>修改密码</b> ... ");
	var sCountUrl = Request.form('CountUrl')+'';
	sCountUrl = sCountUrl.replace(/\\/,"\\\\").replace(/\"/g,'\\\"');
	if(!sCountUrl){
		Response.write("未做修改。<br>");
	}else{
		if(sCountUrl.substr(0,7)=='http://'&&sCountUrl.substr(sCountUrl.length-1)=='/'){
			var s = read("../count.js");
			s=s.replace(/var.+?__cc_sys_url.+?=.+?\".+?\"[ \t;]*$/gim,"var __cc_sys_url = \""+sCountUrl+"\"");
			if(write("../count.js",s)) Response.write("操作失败，原因：你的服务器禁止了FSO组件。<br>");
			else Response.write("完成。<br>")
		}else{
			Response.write("操作失败：路径有错误，请以http://开头，以/结尾。<br>");
		}
	}
	Response.flush();

	Response.write("<br><b>保存设置</b> ... ");
	var sTimeZoneOffset = parseInt(Request.form('TimeZoneOffset')+'');
	var sAuthorize = Request.form('Authorize')+'';
	var sWebMasterEmail = Request.form('WebMasterEmail')+'';
	var sAllowRegister = Request.form('AllowRegister')+'';
	var sDbPath = Request.form('DbPath')+'';
	if(sAllowRegister!='true') sAllowRegister='false';
	sAuthorize = sAuthorize.replace(/\\/,"\\\\").replace(/\"/g,'\\\"');
	sWebMasterEmail = sWebMasterEmail.replace(/\\/,"\\\\").replace(/\"/g,'\\\"');
	sDbPath = sDbPath.replace(/\\/,"\\\\").replace(/\"/g,'\\\"');
	var s = read("../_inc/common.asp");
	s=s.replace(/Accredit.+?=.+?\".+?\"[ \t]*$/gim,"Accredit			=	\""+sAuthorize+"\"");
	s=s.replace(/WebMasterEmail.+?=.+?\".+?\"[ \t]*$/gim,"WebMasterEmail		=	\""+sWebMasterEmail+"\"");
	s=s.replace(/AllowRegister.+?=.+?(true|false)[ \t]*$/gim,"AllowRegister		=	"+sAllowRegister+"");
	s=s.replace(/TimeZoneOffset.+?=.+?\d+?[ \t]*$/gim,"TimeZoneOffset		=	"+sTimeZoneOffset+"");
	s=s.replace(/DbPath.+?=.+?\".+?\"[ \t]*$/gim,"DbPath				=	\""+sDbPath+"\"");
	if(write("../_inc/common.asp",s)) Response.write("操作失败，原因：你的服务器禁止了FSO组件。<br>");
	else Response.write("完成。<br>");
	
	Response.flush();
	
	Response.write("<br><b>清理IP库</b> ... ");
	if(Request.form("ClearIpDb").Count>0){
		oConn.open();
		oConn.execute("delete from t_ip_package where country like '%未知%'");
		oConn.execute("update t_ip_package set city=' ' where city like'%http://www.wzpg.com%'");
		oConn.close();
		Response.write("完成。<br>");
	}else{
		Response.write("没有操作。<br>");
	}
%>
</div>
</fieldset>
<!--#include file="toolbar.asp"-->
</form>
</body>
</html>