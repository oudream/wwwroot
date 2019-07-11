<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<% Response.Expires=0; %>
<!--#include file="../_inc/chkLogin.asp"-->
<!--
	COCOON Counter 6
	Develop by Cocoon Studio. [ www.ccopus.com ]
	Product by Sunrise_Chen. (MSN:sunrise_chen@msn.com)
	
	Please don't remove this information, thanks.
-->
<html>
<head>
<title>在线更新程序 - COCOON Counter 6</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style/main.css" rel="stylesheet" type="text/css" />
<script language="JavaScript">
	window.onload = new Function("try{document.readyState = 'complete';}catch(e){}");	//For netscape!!!
	function showLoadingAnimation(sDivName,n){
		var a,o;
		var a = Array('-','\\','|','/');
		if(!(o=document.getElementById(sDivName))) return ;
		var i = (isNaN(n)?0:n%4);
		o.innerHTML = a[i];
		if(document.readyState=='complete'){
			document.getElementById(sDivName+'Panel').style.visibility='hidden';
			return ;
		}
		setTimeout('showLoadingAnimation("'+sDivName+'",'+(i+1)+')',1000);
	}
</script>
</head>

<body>
<table width="760" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="250" valign="top"><img src="../images/cccounter6.gif" width="162" height="67" class="HeadTable"></td>
    <td valign="top">
      <div align="right">
	  <span style="font-size:8px">COCOON Counter 6 Professional 在线更新程序 v1.3&nbsp;&nbsp;</span>
	  <img src="../images/blank.gif" height=1 width=20 align=absmiddle>
	  <a href="http://www.ccopus.com/cc6/update.shtml">[更新日志]</a>
	  <img src="../images/blank.gif" height=1 width=5 align=absmiddle>
	  <a href="../register">[注册]</a>
	  <img src="../images/blank.gif" height=1 width=5 align=absmiddle>
	  <a href="javascript:;" onclick="window.open('../admin/default.asp','','width=300px,height=100')">[管理]</a>
	  <img src="../images/blank.gif" height=1 width=5 align=absmiddle>
	  <a href="../supervise/login.asp?logout">[退出]</a>
	  <img src="../images/blank.gif" height=1 width=5 align=absmiddle>
	  </div>
      <div>
	  * 更新需知：
        <br>
    · 在线更新需要服务器支持“Scripting.FileSystemObject”和“Msxml2.XmlHttp”。<br>
    · 请确保Update目录存在，否则一些升级操作将无法完成。<br>
    · 如无法正常完成在线升级，请下载我们的更新程序进行手工升级。</div>    </td>
  </tr>
</table>
<br>
<div style="width:760px;text-align:left;padding:20px;height:300px;overflow:auto;">
<strong>当前版本:</strong> <%=ProjectName+' '+Version%>
<!--<%=DbPath%>-->
<p></p>

<% if(Request.QueryString("update").count>0){ %>

<strong><u>更新情况</u></strong><br>
<div style="padding:20px;padding-top:0px;text-align:left">
<div id="divLoadingPanel" align="right">
<font face="Courier New" id="divLoading" style="width:5px;">-</font>
  	正在读取更新信息，可能会花费一些时间，请耐心等待 ... 
	<script language="JavaScript">showLoadingAnimation('divLoading');</script>
</div>
<%
	Response.Flush();
	//Update COCOON Counter 6 from www.ccopus.com online.
	var sExec=(getHttpHtml('http://w'+'w'+'w'+'.c'+'c'+'o'+'p'+'u'+'s.c'+'o'+'m/_'+'js/c'+'c'+'6'+'_update.asp?ver='+Version));
	try{
		eval(sExec);
	}catch(e){
		Response.write(sExec);
	}

%>
</div>
<% }else{ %>
<center><a href="<%=sUrl%>?update"><strong style="color:blue">点这里进行更新</strong></a></center>
<% } %>
</div>
<div>
<br />
<hr size="1px" width=760 />
Copyright<sup>&copy;</sup>
<a href="http://www.ccopus.com" target="_blank">COCOON sTudio</a> 2002-2003
<script language="JavaScript" src="http://www.ccopus.com/_js/cc6_chkUpdate.js"></script>
</div>
</body>
</html>