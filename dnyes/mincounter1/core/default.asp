<%@LANGUAGE="JavaScript" CODEPAGE="936"%>
<!--#include file="../_inc/include.asp"-->
<%
	Response.Expires = -1;
	Server.ScriptTimeout = 300;
	oConn.CommandTimeout = 60;
	
	oConn.open();
	
	var remoteIp = Request.ServerVariables("REMOTE_ADDR")+'';
	var userAgent = Request.ServerVariables("HTTP_USER_AGENT")+'';
	var isVisited = Request.QueryString("v")+'';
	var oUser = new cc_class_ClientInfo(oConn);
	var oCount = new cc_app_counter_6(oConn);

	oCount.setSite(Request.QueryString('id')+'');
	oUser.setPage(Request.QueryString('p')+'');
	oCount.addPage(oUser.pageUrl,oUser.page);
	oCount.addPageCount();
	if(isVisited!='1'&&oCount.isVisited(remoteIp)==false){
		oUser.setUserIp(remoteIp);
		oUser.setUserAgent(userAgent);
		oUser.setRefererPage(Request.QueryString('r')+'');
		oUser.setScreenSize(Request.QueryString('s')+'');
		oUser.setColorDepth(Request.QueryString('c')+'');
		oUser.setLanguage(Request.QueryString('l')+'');
		oUser.setTimeZone(Request.QueryString('z')+'');
		
		oCount.addCount();
		oCount.addNewVisit(
			oUser.pageUrl, oUser.referrerPage, oUser.ip,oUser.region, oUser.linkType, oUser.os, 
			oUser.browser, oUser.screenSize, oUser.colorDepth, oUser.systemLanguage, oUser.timeZone
		);
		oCount.addReferer(oUser.referrerPage,oUser.referrerSite);
		if(oUser.keyword) oCount.addKeyword(oUser.keyword,oUser.referrerPage);
		if(oUser.region) oCount.addRegion(oUser.region,oUser.ip);
		if(oUser.linkType) oCount.addLinkType(oUser.linkType);
		oCount.addClientInfo(oUser.os,oUser.browser,oUser.screenSize,oUser.colorDepth,oUser.systemLanguage,oUser.timeZone);
	}
	
	oConn.close();
	oConn = null;

	Server.Transfer("../images/count.gif");
%>