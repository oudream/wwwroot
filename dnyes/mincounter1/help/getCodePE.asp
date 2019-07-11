<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<%
	var sCountUrl = '';
	var sSiteUrl = Request.ServerVariables("SERVER_NAME")(1);
	var sSitePort = Request.ServerVariables("SERVER_PORT")(1);
	if(isNaN(sSitePort)||sSitePort=="80") sSitePort=""; else sSitePort=":"+sSitePort;
	var sCountPath = Request.ServerVariables("URL")(1);
	sCountPath = sCountPath.substr(0,sCountPath.lastIndexOf("/")+1);
	sCountUrl = "http://" + sSiteUrl + sSitePort + sCountPath
	var sCountSript = sCountUrl + "count.js"
%>
<style type="text/css">
<!--
.style2 {
	color: #990000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
}
.style3 {color: #0000FF}
.style4 {
	color: #000066;
	font-weight: bold;
}
.style8 {font-family: Verdana, Arial, Helvetica, sans-serif}
.style9 {color: #990000}
.style10 {color: #990066}
.style11 {color: #339999}
-->
</style>

<span class="style2">&lt;script language=<span class="style3">"JavaScript"</span> src=<span class="style3">"<span class="style8"><%=sCountSript%></span>"</span>&gt;</span><span class="style8"><span class="style9">&lt;/script&gt;</span></span>