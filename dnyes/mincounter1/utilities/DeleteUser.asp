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
<title>������ - COCOON Counter 6</title>
<style>
body,form,legend,td,fieldset {font-size:9pt;cursor:default}
input {border-width:1px;font-size:9pt;}
</style>
<script language=javascript src="../script/nomenu.js"></script>
</head>
<body style="background-color:buttonface;padding:0;margin:5;border:0">
<form method="post" style="margin:0">
<fieldset style="width:100%;height=75px">
<legend>ɾ���û�</legend>
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
		Response.write("�û�������")
		Response.end()
	}
	var sId = oRs.fields.item(0).value
	oConn.execute("delete from t_Visit where SiteId=" + sId)
	Response.write("ɾ����������Ϣ ... ���.<br>\n")
	Response.flush()
	oConn.execute("delete from t_Region where SiteId=" + sId)
	Response.write("ɾ������ͳ����Ϣ ... ���.<br>\n")
	Response.flush()
	oConn.execute("delete from t_Page where SiteId=" + sId)
	Response.write("ɾ��ҳ��ͳ����Ϣ ... ���.<br>\n")
	Response.flush()
	oConn.execute("delete from t_Link where SiteId=" + sId)
	Response.write("ɾ��������ʽͳ����Ϣ ... ���.<br>\n")
	Response.flush()
	oConn.execute("delete from t_RefPage where SiteId=" + sId)
	Response.write("ɾ������ҳͳ����Ϣ ... ���.<br>\n")
	Response.flush()
	oConn.execute("delete from t_Keyword where SiteId=" + sId)
	Response.write("ɾ���ؼ���ͳ����Ϣ ... ���.<br>\n")
	Response.flush()
	oConn.execute("delete from t_Client where SiteId=" + sId)
	Response.write("ɾ���ͻ���ͳ����Ϣ ... ���.<br>\n")
	Response.flush()
	oConn.execute("delete from t_Count where SiteId=" + sId)
	Response.write("ɾ��ʱ��ͳ����Ϣ ... ���.<br>\n")
	Response.flush()
	oConn.execute("delete from t_Site where SiteId=" + sId)
	Response.write("ɾ��������Ϣ ... ���.<br>\n")
	Response.flush()
	oConn.Close()
/*
	Response.write("ѹ�����ݿ� ... ")
	try{
		Server.execute("Compress.asp")
		Response.write("���.<br>\n")
	}catch(e){
		Response.write("ʧ��: <font color=red>"+e.description+"</font>.<br>\n")
	}
*/
%>
</td></tr></table>
</div>
</fieldset>
</form>
</body>
</html>