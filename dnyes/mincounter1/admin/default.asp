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
label {cursor:hand;}
.content{padding-left:20px;}
</style>
<script language=javascript src="../script/nomenu.js"></script>
</head>
<body style="background-color:buttonface;padding:0;margin:5;border:0;overflow:auto" scroll=auto>
<form method="post" style="margin:0;width:100%;height:84%" action="proc.asp">
<fieldset style="width:100%;height:100%">
<legend>数据压缩工具</legend>
<div style="width:100%;height:95%;padding-left:10px;overflow:auto;">
	  <div>
	  	<input type="checkbox" name="CompressPage" id="CompressPage"><label for="CompressPage"><strong>压缩页面访问数据</strong></label>
		<div class="content">删除浏览量小于 5 次的页面访问信息。</div>
	  </div><br>
	  <div>
	  	<input type="checkbox" name="CompressReferrer" id="CompressReferrer"><label for="CompressReferrer"><strong>压缩引用页数据</strong></label>
	  	<div class="content">删除连入次数小于 5 次的站点信息。</div>
	  </div><br>
	  <div>
	  	<input type="checkbox" name="CompressKeyword" id="CompressKeyword"><label for="CompressKeyword"><strong>压缩引用关键字数据</strong></label>
	  	<div class="content">删除引用次数小于 5 次的关键字信息。</div>
	  </div><br>
	  <div>
	  	<input type="checkbox" name="CompressRegion" id="CompressRegion"><label for="CompressRegion"><strong>压缩访问地区数据</strong></label>
	  	<div class="content">删除访问次数小于 5 次的地区信息。</div>
	  </div><br>
	  <div>
	  	<input type="checkbox" name="CompressClient" id="CompressClient"><label for="CompressClient"><strong>压缩客户端数据</strong></label>
	  	<div class="content">删除访问次数小于 5 次的客户端信息。</div>
	  </div><br>
	  <div>
	  	<input type="checkbox" name="FixDB" id="FixDB"><label for="FixDB"><strong>修复错误数据 *</strong></label>
	  	<div class="content">安全修复在使用过程中因某些原因而造成数据重复、数据丢失等错误.
		<br>对正常的统计数据没有影响，请放心使用。</div>
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
	  	<input type="checkbox" name="CompressDB" id="CompressDB" checked><label for="CompressDB"><strong>压缩数据库文件</strong></label>
	  	<%=bError?" - <i>您不能进行此项操作，因为您的服务器不支持所需的组件</i>":""%>
		<div class="content">压缩与修复ACCESS数据库。<br>
	    附加说明: 服务器必须支持FileSystemObject(FSO)。
		
		</div>
	  </div><br>
	  <div align="right">
	  <input type="submit" name="Submit" value="执行" style="width:75px;">&nbsp;&nbsp;&nbsp;&nbsp;
	  </div>
</div>
</fieldset>
<!--#include file="toolbar.asp"-->
</form>
</body>
</html>