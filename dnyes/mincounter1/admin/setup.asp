<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<% chkAdmin() %>
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
<script language="javascript">
document.onkeydown = null;
</script>
</head>
<body style="background-color:buttonface;padding:0;margin:5;border:0;overflow:auto" scroll=auto>
<form style="margin:0;width:100%;height:84%" method="post" action="procsetup.asp">
<fieldset style="width:100%;height:100%">
<legend>��������</legend>
<div style="width:100%;height:95%;padding-left:10px;overflow:auto;">
	<div id="divFunc">
	<br>
		<strong>�޸Ĺ���Ա����</strong><br>
		ԭʼ���룺<input type="password" name="OldPwd"> ��ʼ���룺admin<br>
		�µ����룺<input type="password" name="NewPwd">
		ȷ�����룺<input type="password" name="CfmPwd"><br>
	</div>
	<div id="divFunc">
	<br>
		<strong>���ݿ��������</strong><br>
		���ݿ�·����<input type="text" name="dbPath" value="<%=DbPath%>" size="40"> <br>
	</div>
	<div id="divFunc">
	<br>
		<strong>�޸�ͳ��·��</strong><br>
<%
	var sContent = read("../count.js");
	if(!sContent){
		Response.write("���󣺶�ȡ�ļ�ʧ�ܣ�������ķ�������ֹ��FSO���<br>");
	}
	var a = sContent.match(/var.+?__cc_sys_url.+?=.+?\"(.+?)\"/g);
	if(!a){
		Response.write("����δ�ҵ����ݣ��ļ������ѱ�����<br>");
	}else{
		var b = a[a.length-1].match(/\"(.+?)\"/g);
		Response.Write("��ǰ���ã� "+ b + "<br>");
	}
%>
		�ο����ã� "http://www.ccopus.com/cccounter6/"
		<BR>
		ͳ��·���� <input type="test" name="CountUrl" size="45">
		<br>
	</div>
	<div id="divFunc">
	<br>
	  	<strong>�޸Ļ�������</strong><br>
		�������䣺 <input type="test" name="WebMasterEmail" size="30" value="<%=WebMasterEmail%>"><br>
		��Ȩ��Ϣ�� <input type="test" name="Authorize" size="30" value="<%=Accredit%>"><br>
		ʱ�����ã� <input type="test" name="TimeZoneOffset" size="5" value="<%=TimeZoneOffset%>"> *������һ���������磺-1 �� 8<br>
		ע�᣺<input type="radio" id="regYes" name="AllowRegister" value="true" <%=AllowRegister?'checked':''%>><label for="regYes">����</label>
		<input type="radio" id="regNo" name="AllowRegister" value="false" <%=AllowRegister?'':'checked'%>><label for="regNo">��ֹ</label><br>
	</div>
	<div id="divFunc">
	<br>
	  	<input type="checkbox" name="ClearIpDb" id="ClearIpDb"><label for="ClearIpDb"><strong>����IP��</strong> ɾ��δ֪��IP��ַ���Լӿ�����ٶȡ�(ֻ���ڰ�װ�����IP��ʱִ��)</label>
	</div>
	<br>
	  <div align="right">
	  <input type="submit" value="����" style="width:75px;">
	  <input type="reset" value="ȡ��" style="width:75px;">&nbsp;&nbsp;&nbsp;&nbsp;
	  </div>
	

</div>
</fieldset>
<!--#include file="toolbar.asp"-->
</form>
</body>
</html>