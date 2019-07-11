<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<%
	oConn.Open();
	
	var bIsLogined = chkLogin(true);
	
	var iWidth = 260;
	var aRs = new Array();
	var aClientLink = new Array();
	var iLinkSum=0, iLinkMax=0,iTmpValue=0;
	
	sSqlString = "select Link,LinkCount,CreateTime from t_Link where SiteId=" + iSiteId + ' order by LinkCount desc';
	var oRs = oConn.Execute( sSqlString );
	if(!oRs.EOF){
		var aRs = oRs.GetRows(20);
		for(var i=0;i<=aRs.ubound(2);++i){
			iTmpValue = aRs.getItem(1,i);
			iLinkSum += iTmpValue;
			iLinkMax = (iTmpValue>iLinkMax?iTmpValue:iLinkMax);
			aClientLink[aClientLink.length] = Array(aRs.getItem(0,i),iTmpValue);
		}
	}
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
<title>连接方式统计 - <%=WebTitle%></title>
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
	<table width="760" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="45%" align="left" valign="top">
      <p>*说明： 所有数据皆为已知数据，未知数据不在统计之列。</p>
      <p>目前支持的连接方式包括：拨号接入，
        长城宽带，
        有线通，
        ADSL，专线，
        广电，铁通，
        网通，
        联通，
        吉通，
        电通，
        东方网景，CHINANET，网吧，
        大学，
        电信 等。 </p>
      <p>数据来自“五洲编程 wzpg.com”，感谢他们。</p>
      <p>COCOON Counter 将不定期更新数据库以保证信息的准确性，如果你拥有COCOON Counter的拷贝，则请定期到我们网站上来查询是否有新的更新。</p>
      <p align="right">-- Sunrise_Chen </p></td>
      <td width="55%" align="right" valign="top">
        <table width="400" border="0" cellpadding="2" cellspacing="1" class="OuterTable">
          <tr>
            <td class="OuterHead"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
                <tr>
                  <td nowrap="nowrap" style="text-align:left">
					<span id="ItemTitle"><font face="webdings">8</font> 连接方式统计</span>
				  </td>
                  <td align="right" class="CodingInfo">Coding by Sunrise_Chen. </td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td><table width="100%"  border="0" cellpadding="0" cellspacing="0" class="OuterTable">
    <%
			for(var i=0;i<aClientLink.length;++i){
				iTmpValue = aClientLink[i][1];
				var iWidthPersent = Math.round(iTmpValue/iLinkSum*(iWidth-5)*100)/100+1;
				var iTotalPersent = Math.round(iTmpValue/iLinkSum*10000)/100;
				Response.Write(
					  '<tr>'
					+ '<td class="InnerHead" height="18" style="padding-left:20px;text-align:left">'+aClientLink[i][0]+'</td>'
					+ '<td width="'+iWidth+'" class="InnerMain" style="text-align:left" title="'+aClientLink[i][0]+'\n'+iTmpValue+' ('+iTotalPersent+'%)">'
					+ '<span class="'+(iTmpValue==iLinkMax?'HotColumnH':'ColumnH')+'" style="width:'+iWidthPersent+'px">'+(iTotalPersent>10?iTmpValue+'&nbsp;':'')+'</span>&nbsp;'
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
