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
<title>������ - COCOON Counter 6</title>
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
<legend>������</legend>
<div style="width:100%;height:90%;padding-left:10px;overflow:auto;">
<%
	var sOldPwd = Request.form("OldPwd")+'';
	var sNewPwd = Request.form("NewPwd")+'';
	var sCfmPwd = Request.form("CfmPwd")+'';
	Response.write("<br><b>�޸�����</b> ... ");
	if(!sOldPwd){
		Response.write("δ����.<br>");
	}else{
		var s = read("../_inc/common.asp");
		//sOldPwd = sOldPwd.replace(/\\/,"\\\\").replace(/\"/g,'\\\"');
		if(sOldPwd!=AdminPwd){Response.write("����ʧ�ܣ�ԭ��ԭʼ���벻��ȷ��<br>")}
		else{
			if(sNewPwd!=sCfmPwd){Response.write("����ʧ�ܣ�ԭ��������������벻��ͬ��<br>")}
			else{
				var re = new RegExp('AdminPwd.+?=.+?".+?"[ \t]*$','gm');
				sNewPwd = sNewPwd.replace(/\\/,"\\\\").replace(/\"/g,'\\\"');
				s=s.replace(re,'AdminPwd			=	"'+sNewPwd+'"');
				if(write("../_inc/common.asp",s)) Response.write("����ʧ�ܣ�ԭ����ķ�������ֹ��FSO�����<br>");
				else Response.write("��ɡ�<br>")
			}
		}
	}

	Response.write("<br><b>�޸�����</b> ... ");
	var sCountUrl = Request.form('CountUrl')+'';
	sCountUrl = sCountUrl.replace(/\\/,"\\\\").replace(/\"/g,'\\\"');
	if(!sCountUrl){
		Response.write("δ���޸ġ�<br>");
	}else{
		if(sCountUrl.substr(0,7)=='http://'&&sCountUrl.substr(sCountUrl.length-1)=='/'){
			var s = read("../count.js");
			s=s.replace(/var.+?__cc_sys_url.+?=.+?\".+?\"[ \t;]*$/gim,"var __cc_sys_url = \""+sCountUrl+"\"");
			if(write("../count.js",s)) Response.write("����ʧ�ܣ�ԭ����ķ�������ֹ��FSO�����<br>");
			else Response.write("��ɡ�<br>")
		}else{
			Response.write("����ʧ�ܣ�·���д�������http://��ͷ����/��β��<br>");
		}
	}
	Response.flush();

	Response.write("<br><b>��������</b> ... ");
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
	if(write("../_inc/common.asp",s)) Response.write("����ʧ�ܣ�ԭ����ķ�������ֹ��FSO�����<br>");
	else Response.write("��ɡ�<br>");
	
	Response.flush();
	
	Response.write("<br><b>����IP��</b> ... ");
	if(Request.form("ClearIpDb").Count>0){
		oConn.open();
		oConn.execute("delete from t_ip_package where country like '%δ֪%'");
		oConn.execute("update t_ip_package set city=' ' where city like'%http://www.wzpg.com%'");
		oConn.close();
		Response.write("��ɡ�<br>");
	}else{
		Response.write("û�в�����<br>");
	}
%>
</div>
</fieldset>
<!--#include file="toolbar.asp"-->
</form>
</body>
</html>