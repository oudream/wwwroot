<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<%
	oConn.Open();
	
	var bIsLogined = chkLogin(true);
	
	var iWidth = 210;
	var aClientScreenSize = new Array();
	var iScreenSizeSum=0, iScreenSizeMax=0,iTmpValue=0;
	var aClientColorDeep = new Array();
	var iColorDeepSum=0, iColorDeepMax=0,iTmpValue=0;
	var aCommon = new Array();
	
	sSqlString = "SELECT ScreenSize,sCount FROM "
						 + "("
						 + "	select ScreenSize,Sum(ClientCount) As sCount from t_Client where SiteId=" + iSiteId + " group by ScreenSize "
						 + ") ORDER BY sCount DESC";
	var oRs = oConn.Execute( sSqlString );
	if(!oRs.EOF){
		var aRs = oRs.GetRows();
		for(var i=0;i<=aRs.ubound(2);++i){
			iTmpValue = aRs.getItem(1,i);
			iScreenSizeSum += iTmpValue;
			iScreenSizeMax = (iTmpValue>iScreenSizeMax?iTmpValue:iScreenSizeMax);
			aClientScreenSize[aClientScreenSize.length] = Array(aRs.getItem(0,i),iTmpValue);
		}
	}
	
	sSqlString = "SELECT ColorDeep,sCount FROM "
						 + "("
						 + "	select ColorDeep,Sum(ClientCount) As sCount from t_Client where SiteId=" + iSiteId + " group by ColorDeep "
						 + ") ORDER BY sCount DESC";
	var oRs = oConn.Execute( sSqlString );
	if(!oRs.EOF){
		var aRs = oRs.GetRows();
		for(var i=0;i<=aRs.ubound(2);++i){
			iTmpValue = aRs.getItem(1,i);
			iColorDeepSum += iTmpValue;
			iColorDeepMax = (iTmpValue>iScreenSizeMax?iTmpValue:iScreenSizeMax);
			aClientColorDeep[aClientColorDeep.length] = Array(aRs.getItem(0,i),iTmpValue);
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
<title>客户端屏幕统计 - <%=WebTitle%></title>
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
    <td width="50%" align="left" valign="top"><table width="370" border="0" cellpadding="2" cellspacing="1" class="OuterTable">
      <tr>
        <td class="OuterHead"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
            <tr>
              <td nowrap="nowrap" style="text-align:left">
				  <span id="ItemTitle"><font face="webdings">8</font> 客户端屏幕颜色深度</span>
			  </td>
              <td align="right" class="CodingInfo">Coding by Sunrise_Chen. </td>
            </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%"  border="0" cellpadding="0" cellspacing="0" class="OuterTable">
            <%
			for(var i=0;i<aClientColorDeep.length;++i){
				iTmpValue = aClientColorDeep[i][1];
				var iWidthPersent = Math.round(iTmpValue/iColorDeepSum*(iWidth-5)*100)/100+1;
				var iTotalPersent = Math.round(iTmpValue/iColorDeepSum*10000)/100;
				Response.Write(
					  '<tr>'
					+ '<td class="InnerHead" height="18" style="padding-left:20px;text-align:left">'+aClientColorDeep[i][0]+'</td>'
					+ '<td width="'+iWidth+'" class="InnerMain" style="text-align:left" title="'+aClientColorDeep[i][0]+'\n'+iTmpValue+' ('+iTotalPersent+'%)">'
					+ '<span class="'+(iTmpValue==iColorDeepMax?'HotColumnH':'ColumnH')+'" style="width:'+iWidthPersent+'px">'+iTmpValue+'&nbsp;</span> '
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
    <td width="50%" align="right" valign="top"><table width="370" border="0" cellpadding="2" cellspacing="1" class="OuterTable">
      <tr>
        <td class="OuterHead"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
            <tr>
              <td nowrap="nowrap" style="text-align:left">
				  <span id="ItemTitle"><font face="webdings">8</font> 客户端屏幕大小</span>
			  </td>
              <td align="right" class="CodingInfo">Coding by Sunrise_Chen. </td>
            </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%"  border="0" cellpadding="0" cellspacing="0" class="OuterTable">
            <%
			for(var i=0;i<aClientScreenSize.length;++i){
				iTmpValue = aClientScreenSize[i][1];
				var iWidthPersent = Math.round(iTmpValue/iScreenSizeSum*(iWidth-5)*100)/100+1;
				var iTotalPersent = Math.round(iTmpValue/iScreenSizeSum*10000)/100;
				Response.Write(
					  '<tr>'
					+ '<td class="InnerHead" height="18" style="padding-left:20px;text-align:left">'+aClientScreenSize[i][0].replace('_','×')+'</td>'
					+ '<td width="'+iWidth+'" class="InnerMain" style="text-align:left" title="'+aClientScreenSize[i][0]+'\n'+iTmpValue+' ('+iTotalPersent+'%)">'
					+ '<span class="'+(iTmpValue==iScreenSizeMax?'HotColumnH':'ColumnH')+'" style="width:'+iWidthPersent+'px">'+iTmpValue+'&nbsp;</span> '
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
  </tr>
</table>
<!--#include file="../static/bott.htm"-->
</body>
</html>
