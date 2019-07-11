<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<!--#include file="inc_Language.asp"-->
<%
	oConn.Open();
	
	var bIsLogined = chkLogin(true);
	
	var iWidth = 210;
	var aClientOs = new Array();
	var aClientBrowser = new Array();
	var aCommon = new Array();
	var iOsSum=0, iOsMax=0,iTmpValue=0;
	var iBrowserSum=0, iBrowserMax=0;
	
	sSqlString = "SELECT OS,OSCount FROM "
						 + "("
						 + "	select OS,Sum(ClientCount) As OSCount from t_Client where SiteId=" + iSiteId + " group by OS "
						 + ") a ORDER BY OSCount DESC";
	var oRs = oConn.Execute( sSqlString );
	if(!oRs.EOF){
		var aRs = oRs.GetRows();
		for(var i=0;i<=aRs.ubound(2);++i){
			iTmpValue = aRs.getItem(1,i);
			iOsSum += iTmpValue;
			iOsMax = (iTmpValue>iOsMax?iTmpValue:iOsMax);
			aClientOs[aClientOs.length] = Array(aRs.getItem(0,i),iTmpValue);
		}
	}
	
	sSqlString = "SELECT Browser,BrowserCount FROM "
						 + "("
						 + "	select Browser,Sum(ClientCount) As BrowserCount from t_Client where SiteId=" + iSiteId + " group by Browser "
						 + ") a ORDER BY BrowserCount DESC";
	var oRs = oConn.Execute( sSqlString );
	if(!oRs.EOF){
		var aRs = oRs.GetRows();
		for(var i=0;i<=aRs.ubound(2);++i){
			iTmpValue = aRs.getItem(1,i);
			iBrowserSum += iTmpValue;
			iBrowserMax = (iTmpValue>iBrowserMax?iTmpValue:iBrowserMax);
			aClientBrowser[aClientBrowser.length] = Array(aRs.getItem(0,i),iTmpValue);
		}
	}
	
	sSqlString = "select top 3 OS,Browser,ScreenSize,ColorDeep,ClientCount from t_Client where SiteId=" + iSiteId + " order by ClientCount desc";
	var oRs = oConn.Execute( sSqlString );
	if(!oRs.EOF) aCommon = oRs.GetRows(3);
	
	oConn.Close();
	
	
%>
<!--
	COCOON Counter 6
	Develop by Cocoon Studio. [ www.ccopus.com ]
	Product by Sunrise_Chen. (MSN:sunrise_chen@msn.com)
	
	Please don't remove this information, thanks.
-->
<html>
<head>
<title>客户端软件统计 - <%=WebTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="../style/main.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.box{border:1px solid #C0C0C0;width:52px;height:19px;clip:rect(0px,181px,18px,0px);overflow:hidden;}
.box2{border:1px solid #F4F4F4;width:50px;height:17px;clip:rect(0px,179px,16px,0px);overflow:hidden;}
select{position:relative;left:-2px;top:-2px;font-size:12px;width:53px;line-height:14px;border:0px;color:#909993;}
-->
</style>
</head>

<body>
<!--#include file="inc_head.asp"-->
	<table width="760" border="0" cellpadding="2" cellspacing="1" class="OuterTable">
    <tr>
      <td class="OuterHead"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
          <tr>
            <td nowrap="nowrap" style="text-align:left">
				<span id="ItemTitle"><font face="webdings">8</font> 常见客户端配置</span>
			</td>
            <td align="right" class="CodingInfo">Coding by Sunrise_Chen. </td>
          </tr>
      </table></td>
    </tr>
    <tr>
      <td><table width="100%"  border="0" cellpadding="0" cellspacing="0" class="OuterTable">
				<tr>
					<td class="InnerHead" width="25%">操作系统</td>
					<td class="InnerHead" width="20%">浏览器</td>
					<td class="InnerHead" width="20%">屏幕大小</td>
					<td class="InnerHead" width="20%">颜色深度</td>
					<td class="InnerHead" width="15%">统计</td>
				</tr>
     <%
		try{
			for(var i=0;i<=aCommon.ubound(2);++i){
				Response.Write(
					  '<tr align="center">'
					+ '<td class="InnerMain" style="text-align:left;padding-left:20px">' + aCommon.getItem(0,i) + '</td>'
					+ '<td class="InnerMain">' + aCommon.getItem(1,i) + '</td>'
					+ '<td class="InnerMain">' + aCommon.getItem(2,i) + '</td>'
					+ '<td class="InnerMain">' + aCommon.getItem(3,i) + '</td>'
					+ '<td class="InnerMain">' + aCommon.getItem(4,i) + '</td>'
					+ '</tr>' + '\n'
				);
			}
		}catch(e){if(e.name=='TypeError') Response.Write('<tr><td colspan=4></td></tr>')}
	%>
      </table></td>
    </tr>
    <tr>
      <td class="OuterFoot"></td>
    </tr>
  </table>
	<br>
	<table width="760" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%" align="left" valign="top">
        <table width="370" border="0" cellpadding="2" cellspacing="1" class="OuterTable">
          <tr>
            <td class="OuterHead"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
                <tr>
                  <td nowrap="nowrap" style="text-align:left">
					  <span id="ItemTitle"><font face="webdings">8</font> 客户端浏览器统计</span>
				  </td>
                  <td align="right" class="CodingInfo">Coding by Sunrise_Chen. </td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td><table width="100%"  border="0" cellpadding="0" cellspacing="0" class="OuterTable">
                <%
			for(var i=0;i<aClientBrowser.length;++i){
				iTmpValue = aClientBrowser[i][1];
				var iWidthPersent = Math.round(iTmpValue/iBrowserSum*(iWidth-5)*100)/100+1;
				var iTotalPersent = Math.round(iTmpValue/iBrowserSum*10000)/100;
				Response.Write(
					  '<tr>'
					+ '<td class="InnerHead" height="18" style="padding-left:20px;text-align:left">'+aClientBrowser[i][0]+'</td>'
					+ '<td width="'+iWidth+'" class="InnerMain" style="text-align:left" title="'+aClientBrowser[i][0]+'\n'+iTmpValue+' ('+iTotalPersent+'%)">'
					+ '<span class="'+(iTmpValue==iBrowserMax?'HotColumnH':'ColumnH')+'" style="width:'+iWidthPersent+'px">'+iTmpValue+'&nbsp;</span> '
					+ '<span style="vertical-align:middle;font-size:8pt">('+iTotalPersent+'%)</span>'
					+ '</td>'
					+ '</tr>' + '\n'
				);
			}
		%>
            </table></td>
          </tr>
          <tr>
            <td class="OuterFoot"></td>
          </tr>
        </table>      </td>
      <td align="right" valign="top">
        <table width="370" border="0" cellpadding="2" cellspacing="1" class="OuterTable">
          <tr>
            <td class="OuterHead"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
                <tr>
                  <td nowrap="nowrap" style="text-align:left">
					<span id="ItemTitle"><font face="webdings">8</font> 客户端操作系统统计</span>
				  </td>
                  <td align="right" class="CodingInfo">Coding by Sunrise_Chen. </td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td><table width="100%"  border="0" cellpadding="0" cellspacing="0" class="OuterTable">
    <%
			for(var i=0;i<aClientOs.length;++i){
				iTmpValue = aClientOs[i][1];
				var iWidthPersent = Math.round(iTmpValue/iOsSum*(iWidth-5)*100)/100+1;
				var iTotalPersent = Math.round(iTmpValue/iOsSum*10000)/100;
				Response.Write(
					  '<tr>'
					+ '<td class="InnerHead" height="18" style="padding-left:20px;text-align:left">'+aClientOs[i][0]+'</td>'
					+ '<td width="'+iWidth+'" class="InnerMain" style="text-align:left" title="'+aClientOs[i][0]+'\n'+iTmpValue+' ('+iTotalPersent+'%)">'
					+ '<span class="'+(iTmpValue==iOsMax?'HotColumnH':'ColumnH')+'" style="width:'+iWidthPersent+'px">'+iTmpValue+'&nbsp;</span> '
					+ '<span style="vertical-align:middle;font-size:8pt">('+iTotalPersent+'%)</span>'
					+ '</td>'
					+ '</tr>' + '\n'
				);
			}
		%>
            </table></td>
          </tr>
          <tr>
            <td class="OuterFoot"></td>
          </tr>
        </table>
        </td>
    </tr>
  </table>
<!--#include file="../static/bott.htm"-->
</body>
</html>
