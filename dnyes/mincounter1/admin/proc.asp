<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<%
	iDeleteCount = 5;

	if(Request.ServerVariables("REQUEST_METHOD")!="POST")
		Response.end();
	
	chkAdmin();
%>
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
<form method="post" style="margin:0;width:100%;height:84%">
<fieldset style="width:100%;height:100%">
<legend>管理工具</legend>
<div style="width:100%;height:90%;padding-left:10px;overflow:auto;">
<%
	oConn.Open();

	//压缩页面访问数据
	if(Request.form("CompressPage").Count){
		Response.write("<br>");
		Response.write("压缩页面访问数据 ... ");
		Response.flush();
		var t1 = new Date();
		try{
			sSqlString = "delete from t_Page where PageCount<=" + iDeleteCount ;
			oConn.execute( sSqlString );
			var t2 = new Date();
			Response.write("完成， ");
			Response.write("耗时 "+(t2-t1)/1000 + " 秒。");
		}catch(e){
			Response.write("发生错误; 描述: "+e.description);
		}
		Response.write("<br>");
		Response.flush();
	}

	//压缩引用页数据
	if(Request.form("CompressReferrer").Count){
		Response.write("<br>");
		Response.write("压缩引用页数据 ... ");
		Response.flush();
		var t1 = new Date();
		try{
			sSqlString = "delete from t_RefPage where RefPageCount<=" + iDeleteCount ;
			oConn.execute( sSqlString );
			var t2 = new Date();
			Response.write("完成， ");
			Response.write("耗时 "+(t2-t1)/1000 + " 秒。");
		}catch(e){
			Response.write("发生错误; 描述: "+e.description);
		}
		Response.write("<br>");
		Response.flush();
	}

	//压缩引用关键字数据
	if(Request.form("CompressKeyword").Count){
		Response.write("<br>");
		Response.write("压缩引用关键字数据 ... ");
		Response.flush();
		var t1 = new Date();
		try{
			sSqlString = "delete from t_Keyword where KeywordCount<=" + iDeleteCount ;
			oConn.execute( sSqlString );
			var t2 = new Date();
			Response.write("完成， ");
			Response.write("耗时 "+(t2-t1)/1000 + " 秒。");
		}catch(e){
			Response.write("发生错误; 描述: "+e.description);
		}
		Response.write("<br>");
		Response.flush();
	}

	//压缩访问地区数据
	if(Request.form("CompressRegion").Count){
		Response.write("<br>");
		Response.write("压缩访问地区数据 ... ");
		Response.flush();
		var t1 = new Date();
		try{
			sSqlString = "delete from t_Region where RegionCount<=" + iDeleteCount ;
			oConn.execute( sSqlString );
			var t2 = new Date();
			Response.write("完成， ");
			Response.write("耗时 "+(t2-t1)/1000 + " 秒。");
		}catch(e){
			Response.write("发生错误; 描述: "+e.description);
		}
		Response.write("<br>");
		Response.flush();
	}

	//压缩客户端数据
	if(Request.form("CompressClient").Count){
		Response.write("<br>");
		Response.write("压缩客户端数据 ... ");
		Response.flush();
		var t1 = new Date();
		try{
			sSqlString = "delete from t_Client where ClientCount<=" + iDeleteCount ;
			oConn.execute( sSqlString );
			var t2 = new Date();
			Response.write("完成， ");
			Response.write("耗时 "+(t2-t1)/1000 + " 秒。");
		}catch(e){
			Response.write("发生错误; 描述: "+e.description);
		}
		Response.write("<br>");
		Response.flush();
	}

	//安全修复统计数据
	if(Request.form("FixDB").Count){
		Response.write("<br>");
		Response.write("安全修复统计数据 ... ");
		Response.flush();
		var t1 = new Date();
		try{
			sSqlString = "delete from QueryDuplicateData" ;
			oConn.execute( sSqlString );
			var t2 = new Date();
			Response.write("完成， ");
			Response.write("耗时 "+(t2-t1)/1000 + " 秒。");
		}catch(e){
			Response.write("发生错误; 描述: "+e.description);
		}
		Response.write("<br>");
		Response.flush();

		Response.write("<br>");
		Response.write("恢复Application变量数据 ... ");
		Response.flush();
		var t1 = new Date();
		try{
			sSqlString = "select SiteId,UID from t_Site order by SiteId";
			var oRs = oConn.Execute( sSqlString );
			while(!oRs.EOF){
				var sSiteId = oRs.Fields.Item(0).Value;
				var sSiteUID = oRs.Fields.Item(1).Value;
				Application.Contents("cc_6_site_uid_"+sSiteUID+"") = sSiteId;
				Application.Contents("cc_6_site_"+sSiteId+"_visited") = "";
				oRs.MoveNext();
			}
			var t2 = new Date();
			Response.write("完成， ");
			Response.write("耗时 "+(t2-t1)/1000 + " 秒。");
			var oRs = null;
		}catch(e){
			Response.write("发生错误; 描述: "+e.description);
		}
		Response.write("<br>");
		Response.flush();
	}

	oConn.Close();
	oConn = null;
	
	//压缩数据库
	if(Request.form("CompressDB").Count){
		Response.write("<br>");
		Response.write("压缩数据库 ... ");
		Response.flush();
		var t1 = new Date();
		try{
			var oConn = new ActiveXObject("ADODB.Connection");
			oConn.ConnectionString = sConnString;
			var oFso = new ActiveXObject("Scripting.FileSystemObject");
			var sDbPath = Server.MapPath(DbPath);
			var oFile = oFso.getFile(sDbPath);
			var iBeginSize = oFile.size;
			var iEndSize = 0;
			
			var oJro = new ActiveXObject("jro.JetEngine");	
			var sTmpDbPath = Server.MapPath(DbPath + ".bak");
			var sTmpConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+sTmpDbPath+";";
			if(oFso.FileExists(sTmpDbPath)) oFso.DeleteFile(sTmpDbPath);
			oJro.CompactDatabase(oConn.ConnectionString,sTmpConnString);
			oFso.CopyFile(sTmpDbPath,sDbPath);
			oFso.DeleteFile(sTmpDbPath);
			oFile = oFso.getFile(sDbPath);
			iEndSize = oFile.size;
			oJro = null;
			oFso = null;
			oConn = null;
			var t2 = new Date();
			Response.write("完成， ");
			Response.write("耗时 "+(t2-t1)/1000 + " 秒。");
			Response.write("<BR>数据库大小: 压缩前: "+(Math.round((iBeginSize/1024/1024)*100)/100) + " MB, "
				+ "压缩后:"+(Math.round((iEndSize/1024/1024)*100)/100)+" MB, "
				+ "缩小了 "+(Math.round(((iBeginSize-iEndSize)/1024/1024)*100)/100)+" MB 。")
		}catch(e){
			Response.write("发生错误，描述: "+e.description);
		}
		Response.write("<br>");
		Response.flush();
	}
%>
</div>
</fieldset>
<!--#include file="toolbar.asp"-->
</form>
</body>
</html>