<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<!--#include file="inc_Language.asp"-->
<!--#include file="inc_TimeZone.asp"-->
<%
	oConn.Open();
	
	var bIsLogined = chkLogin(true);
	
	var iWidth = 210;
	var aClientLanguage = new Array();
	var aClientTimeZone = new Array();
	var aCommon = new Array();
	var iTmpValue=0;
	var iLanguageSum=0, iLanguageMax=0;
	var iTimeZoneSum=0, iTimeZoneMax=0;
	
	sSqlString = "SELECT SystemLanguage,LanguageCount FROM "
						 + "("
						 + "	select SystemLanguage,Sum(ClientCount) As LanguageCount from t_Client where SiteId=" + iSiteId + " and len(SystemLanguage)>0 group by SystemLanguage "
						 + ") a ORDER BY LanguageCount DESC";
	var oRs = oConn.Execute( sSqlString );
	if(!oRs.EOF){
		var aRs = oRs.GetRows();
		for(var i=0;i<=aRs.ubound(2);++i){
			iTmpValue = aRs.getItem(1,i);
			iLanguageSum += iTmpValue;
			iLanguageMax = (iTmpValue>iLanguageMax?iTmpValue:iLanguageMax);
			aClientLanguage[aClientLanguage.length] = Array(oLanguage[aRs.getItem(0,i)],iTmpValue);
		}
	}
	
	sSqlString = "SELECT TimeZone,TimeZoneCount FROM "
						 + "("
						 + "	select TimeZone,Sum(ClientCount) As TimeZoneCount from t_Client where SiteId=" + iSiteId + " and len(TimeZone)>0 group by TimeZone "
						 + ") a ORDER BY TimeZoneCount DESC";
	var oRs = oConn.Execute( sSqlString );
	var sTimeZone,sTimeZoneArea
	if(!oRs.EOF){
		var aRs = oRs.GetRows();
		for(var i=0;i<=aRs.ubound(2);++i){
			iTmpValue = aRs.getItem(1,i);
			iTimeZoneSum += iTmpValue;
			iTimeZoneMax = (iTmpValue>iTimeZoneMax?iTmpValue:iTimeZoneMax);
			sTimeZone = oTimeZone[aRs.getItem(0,i)] + ' ';
			sTimeZoneArea = sTimeZone.substr(sTimeZone.indexOf(' ')+1);
			sTimeZone = sTimeZone.substr(0,sTimeZone.indexOf(' '));
			aClientTimeZone[aClientTimeZone.length] = Array(sTimeZone,iTmpValue,sTimeZoneArea);
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
<title>客户端区域统计 - <%=WebTitle%></title>
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
                  <td nowrap="nowrap" style="text-align:left"> <span id="ItemTitle"><font face="webdings">8</font> 客户端语种</span> </td>
                  <td align="right" class="CodingInfo">Coding by Sunrise_Chen. </td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td><table width="100%"  border="0" cellpadding="0" cellspacing="0" class="OuterTable">
                <%
			iWidth = 150;
			for(var i=0;i<aClientLanguage.length;++i){
				iTmpValue = aClientLanguage[i][1];
				var iWidthPersent = Math.round(iTmpValue/iLanguageSum*(iWidth-5)*100)/100+1;
				var iTotalPersent = Math.round(iTmpValue/iLanguageSum*10000)/100;
				Response.Write(
					  '<tr>'
					+ '<td class="InnerHead" height="18" style="padding-left:10px;text-align:left">'+aClientLanguage[i][0]+'</td>'
					+ '<td width="'+iWidth+'" class="InnerMain" style="text-align:left" title="'+aClientLanguage[i][0]+'\n'+iTmpValue+' ('+iTotalPersent+'%)">'
					+ '<span class="'+(iTmpValue==iLanguageMax?'HotColumnH':'ColumnH')+'" style="width:'+iWidthPersent+'px">'+iTmpValue+'&nbsp;</span> '
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
        </table></td>
      <td align="right" valign="top">
        <table width="370" border="0" cellpadding="2" cellspacing="1" class="OuterTable">
          <tr>
            <td class="OuterHead"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
                <tr>
                  <td nowrap="nowrap" style="text-align:left"> <span id="ItemTitle"><font face="webdings">8</font> 客户端时区</span></td>
                  <td align="right" class="CodingInfo">Coding by Sunrise_Chen. </td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td><table width="100%"  border="0" cellpadding="0" cellspacing="0" class="OuterTable">
         <%
			iWidth = 250;
			for(var i=0;i<aClientTimeZone.length;++i){
				iTmpValue = aClientTimeZone[i][1];
				var iWidthPersent = Math.round(iTmpValue/iTimeZoneSum*(iWidth-5)*100)/100+1;
				var iTotalPersent = Math.round(iTmpValue/iTimeZoneSum*10000)/100;
				Response.Write(
					  '<tr>'
					+ '<td class="InnerHead" height="18" style="padding-left:10px;text-align:left">'+aClientTimeZone[i][0]+'</td>'
					+ '<td width="'+iWidth+'" class="InnerMain" style="text-align:left" title="时区: '+aClientTimeZone[i][0]+'\n流量: '+iTmpValue+' ('+iTotalPersent+'%)\n地区: '+aClientTimeZone[i][2]+'">'
					+ '<span class="'+(iTmpValue==iTimeZoneMax?'HotColumnH':'ColumnH')+'" style="width:'+iWidthPersent+'px">'+iTmpValue+'&nbsp;</span> '
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
        <br>
      </td>
    </tr>
  </table>
<!--#include file="../static/bott.htm"-->
</body>
</html>
