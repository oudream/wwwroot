<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<%
	var sCountUrl = '';
	var sCountUID = '&lt;ÄãµÄÍ³¼ÆÕÊºÅ&gt;';
	var sSiteUrl = Request.ServerVariables("SERVER_NAME")(1);
	var sSitePort = Request.ServerVariables("SERVER_PORT")(1);
	if(isNaN(sSitePort)||sSitePort=="80") sSitePort=""; else sSitePort=":"+sSitePort;
	var sCountPath = Request.ServerVariables("URL")(1);
	sCountPath = sCountPath.substr(0,sCountPath.lastIndexOf("/"));
	sCountPath = sCountPath.substr(0,sCountPath.lastIndexOf("/"));
	sCountUrl = "http://" + sSiteUrl + sSitePort + sCountPath + "/";
	var sCountSript = sCountUrl + "count.js";
	
	if(chkLogin(false)){
		oConn.Open();
		var oRs = oConn.execute( "select UID from t_Site where SiteId=" + iSiteId );
		if(!oRs.EOF){
			sCountUID = oRs.fields.item(0).value;
		}
		oConn.Close();
	}
%>
<style type="text/css">
<!--
body {font-size:12px}
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
<nobr>
<span class="style2">&lt;script language=<span class="style3">"JavaScript"</span>&gt;</span><span class="style8">
<BR>
&nbsp;&nbsp;&nbsp;&nbsp;<span class="style4">var</span> __cc_uid=<span class="style3">"<%=sCountUID%>"</span>;
<BR>
<span class="style9">&lt;/script&gt;<span class="style2">&lt;script language=<span class="style3">"JavaScript"</span> 
<BR>&nbsp;&nbsp;&nbsp;&nbsp;src=<span class="style3">"<%=sCountSript%>"</span>&gt;
<BR>&lt;/script&gt;</span></span></span>
</nobr>