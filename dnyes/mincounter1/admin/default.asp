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
</head>
<body style="background-color:buttonface;padding:0;margin:5;border:0;overflow:auto" scroll=auto>
<form method="post" style="margin:0;width:100%;height:84%" action="proc.asp">
<fieldset style="width:100%;height:100%">
<legend>����ѹ������</legend>
<div style="width:100%;height:95%;padding-left:10px;overflow:auto;">
	  <div>
	  	<input type="checkbox" name="CompressPage" id="CompressPage"><label for="CompressPage"><strong>ѹ��ҳ���������</strong></label>
		<div class="content">ɾ�������С�� 5 �ε�ҳ�������Ϣ��</div>
	  </div><br>
	  <div>
	  	<input type="checkbox" name="CompressReferrer" id="CompressReferrer"><label for="CompressReferrer"><strong>ѹ������ҳ����</strong></label>
	  	<div class="content">ɾ���������С�� 5 �ε�վ����Ϣ��</div>
	  </div><br>
	  <div>
	  	<input type="checkbox" name="CompressKeyword" id="CompressKeyword"><label for="CompressKeyword"><strong>ѹ�����ùؼ�������</strong></label>
	  	<div class="content">ɾ�����ô���С�� 5 �εĹؼ�����Ϣ��</div>
	  </div><br>
	  <div>
	  	<input type="checkbox" name="CompressRegion" id="CompressRegion"><label for="CompressRegion"><strong>ѹ�����ʵ�������</strong></label>
	  	<div class="content">ɾ�����ʴ���С�� 5 �εĵ�����Ϣ��</div>
	  </div><br>
	  <div>
	  	<input type="checkbox" name="CompressClient" id="CompressClient"><label for="CompressClient"><strong>ѹ���ͻ�������</strong></label>
	  	<div class="content">ɾ�����ʴ���С�� 5 �εĿͻ�����Ϣ��</div>
	  </div><br>
	  <div>
	  	<input type="checkbox" name="FixDB" id="FixDB"><label for="FixDB"><strong>�޸��������� *</strong></label>
	  	<div class="content">��ȫ�޸���ʹ�ù�������ĳЩԭ�����������ظ������ݶ�ʧ�ȴ���.
		<br>��������ͳ������û��Ӱ�죬�����ʹ�á�</div>
	  </div><br>
<%
	var bError = false;
	try{
		var oFso = new ActiveXObject("Scripting.FileSystemObject");
		var oJro = new ActiveXObject("jro.JetEngine");
		oFso = null;
		oJro = null;
	}catch(e){
		bError = true;
	}
%>
	  <div <%=bError?"disabled":""%>>
	  	<input type="checkbox" name="CompressDB" id="CompressDB" checked><label for="CompressDB"><strong>ѹ�����ݿ��ļ�</strong></label>
	  	<%=bError?" - <i>�����ܽ��д����������Ϊ���ķ�������֧����������</i>":""%>
		<div class="content">ѹ�����޸�ACCESS���ݿ⡣<br>
	    ����˵��: ����������֧��FileSystemObject(FSO)��
		
		</div>
	  </div><br>
	  <div align="right">
	  <input type="submit" name="Submit" value="ִ��" style="width:75px;">&nbsp;&nbsp;&nbsp;&nbsp;
	  </div>
</div>
</fieldset>
<!--#include file="toolbar.asp"-->
</form>
</body>
</html>