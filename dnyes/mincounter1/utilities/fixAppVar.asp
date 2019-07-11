<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<%
	oConn.Open()

	sSqlString = "select SiteId,UID from t_Site order by SiteId";
	var oRs = oConn.Execute( sSqlString );
	while(!oRs.EOF){
		var sSiteId = oRs.Fields.Item(0).Value;
		var sSiteUID = oRs.Fields.Item(1).Value;
		Application.Contents("cc_6_site_uid_"+sSiteUID+"") = sSiteId;
		Application.Contents("cc_6_site_"+sSiteId+"_visited") = "";
		oRs.MoveNext()
	}

	oConn.Close()
%>

<html>
<head>
<title>COCOON Counter 6 修复工具</title>
</head>
<body>
修复成功。
</body>
</html>