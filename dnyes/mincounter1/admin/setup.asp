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
<script language="javascript">
document.onkeydown = null;
</script>
</head>
<body style="background-color:buttonface;padding:0;margin:5;border:0;overflow:auto" scroll=auto>
<form style="margin:0;width:100%;height:84%" method="post" action="procsetup.asp">
<fieldset style="width:100%;height:100%">
<legend>基本设置</legend>
<div style="width:100%;height:95%;padding-left:10px;overflow:auto;">
	<div id="divFunc">
	<br>
		<strong>修改管理员密码</strong><br>
		原始密码：<input type="password" name="OldPwd"> 初始密码：admin<br>
		新的密码：<input type="password" name="NewPwd">
		确认密码：<input type="password" name="CfmPwd"><br>
	</div>
	<div id="divFunc">
	<br>
		<strong>数据库参数设置</strong><br>
		数据库路径：<input type="text" name="dbPath" value="<%=DbPath%>" size="40"> <br>
	</div>
	<div id="divFunc">
	<br>
		<strong>修改统计路径</strong><br>
<%
	var sContent = read("../count.js");
	if(!sContent){
		Response.write("错误：读取文件失败，可能你的服务器禁止了FSO组件<br>");
	}
	var a = sContent.match(/var.+?__cc_sys_url.+?=.+?\"(.+?)\"/g);
	if(!a){
		Response.write("错误：未找到内容，文件可能已被更改<br>");
	}else{
		var b = a[a.length-1].match(/\"(.+?)\"/g);
		Response.Write("当前设置： "+ b + "<br>");
	}
%>
		参考设置： "http://www.ccopus.com/cccounter6/"
		<BR>
		统计路径： <input type="test" name="CountUrl" size="45">
		<br>
	</div>
	<div id="divFunc">
	<br>
	  	<strong>修改基本设置</strong><br>
		管理信箱： <input type="test" name="WebMasterEmail" size="30" value="<%=WebMasterEmail%>"><br>
		授权信息： <input type="test" name="Authorize" size="30" value="<%=Accredit%>"><br>
		时差设置： <input type="test" name="TimeZoneOffset" size="5" value="<%=TimeZoneOffset%>"> *请输入一个整数，如：-1 或 8<br>
		注册：<input type="radio" id="regYes" name="AllowRegister" value="true" <%=AllowRegister?'checked':''%>><label for="regYes">允许</label>
		<input type="radio" id="regNo" name="AllowRegister" value="false" <%=AllowRegister?'':'checked'%>><label for="regNo">禁止</label><br>
	</div>
	<div id="divFunc">
	<br>
	  	<input type="checkbox" name="ClearIpDb" id="ClearIpDb"><label for="ClearIpDb"><strong>清理IP库</strong> 删除未知的IP地址，以加快检索速度。(只需在安装或更新IP库时执行)</label>
	</div>
	<br>
	  <div align="right">
	  <input type="submit" value="保存" style="width:75px;">
	  <input type="reset" value="取消" style="width:75px;">&nbsp;&nbsp;&nbsp;&nbsp;
	  </div>
	

</div>
</fieldset>
<!--#include file="toolbar.asp"-->
</form>
</body>
</html>