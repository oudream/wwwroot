<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<%
	oConn.Open();
	
	var bIsLogined = chkLogin(true);
	
	var iWidth = 170;

	var aClientPage = new Array();
	var iPageSum=0, iPageMax=0,iTmpValue=0;
	
	var iFound = 0;
	var sOrder = '';
	var sOrderBy = '';
	var sCondition = '';
	var sSearchPage = '';
	if(Request.Form('SearchPage').Count>0)	sSearchPage = Request.Form('SearchPage')(1);
	if(Request.QueryString('OrderBy').Count>0) sOrder = Request.QueryString('OrderBy')(1).toLowerCase();
	switch(sOrder){
	case "time desc":
		sOrderBy = "CreateTime DESC"; break;
	case "time asc":
		sOrderBy = "CreateTime ASC"; break;
	case "count desc":
		sOrderBy = "PageCount DESC"; break;
	case "count asc":
		sOrderBy = "PageCount ASC"; break;
	default:
		sOrder = "count desc";
		sOrderBy = "PageCount DESC"; break;
	}
	if(sSearchPage){
		sSqlString = "select top 20 Page,PageCount,CreateTime from t_Page where SiteId=" + iSiteId + ' and Page like \'%'+ sSearchPage +'%\'';
	}else{
		sSqlString = "select top 20 Page,PageCount,CreateTime from t_Page where SiteId=" + iSiteId + ' order by '+sOrderBy;
	}
	var oRs = oConn.Execute( sSqlString );
	if(!oRs.EOF){
		var aRs = oRs.GetRows(20);
		for(var i=0;i<=aRs.ubound(2);++i){
			var sTmpText = aRs.getItem(0,i);
			if(sTmpText.indexOf(sSearchPage)>-1) iFound++;
			iTmpValue = aRs.getItem(1,i);
			var sLastTime = formatDateTime(new Date(aRs.getItem(2,i)),0);
			iPageMax = (iTmpValue>iPageMax?iTmpValue:iPageMax);
			aClientPage[aClientPage.length] = Array(sTmpText,iTmpValue,sLastTime);
		}
			
		sSqlString = "select sum(PageCount) from t_Page where SiteId=" + iSiteId
		iPageSum = oConn.Execute(sSqlString)(0).Value;
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
<title>页面统计 - <%=WebTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="../style/main.css" rel="stylesheet" type="text/css" />
</head>

<body>
<!--#include file="inc_head.asp"-->
	<table width="760" border="0" cellpadding="2" cellspacing="1" class="OuterTable">
    <tr>
      <td class="OuterHead"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
          <tr>
            <td nowrap="nowrap" style="text-align:left;width:150px">
			<span id="ItemTitle"><font face="webdings">8</font> 页面统计 (前20位)</span></td>
            <td nowrap="nowrap" style="text-align:left;width:400px">
				<form name="formSearchPage" method="post" class="SubHeader">
              　搜索
                <input type="text" name="SearchPage" class="ThinBorderText" style="width:130px" value="<%=sSearchPage%>">
                <input name="button" type="button" onClick="form.submit();" value="查询" class="button">
								&nbsp;&nbsp;&nbsp;<font color="red"><%=(sSearchPage?"找到 " + iFound + " 条记录。":"")%></font>
              </form>
			  </td>
            <td align="right" class="CodingInfo">Coding by Sunrise_Chen. </td>
          </tr>
      </table></td>
    </tr>
    <tr>
      <td><table width="100%"  border="0" cellpadding="0" cellspacing="0" class="OuterTable">
				<tr>
					<td class="InnerHead">访问页</td>
					<td class="InnerHead">
						<%
							if(sOrder=="time desc"){
								Response.Write('<a href="'+sUrl+'?OrderBy=time asc">访问时间 ↓</a>');
							}else if(sOrder=="time asc"){
								Response.Write('<a href="'+sUrl+'?OrderBy=time desc">访问时间 ↑</a>');
							}else{
								Response.Write('<a href="'+sUrl+'?OrderBy=time desc">访问时间</a>');
							}
						%>
					</td>
					<td class="InnerHead">
						<%
							if(sOrder=="count desc"){
								Response.Write('<a href="'+sUrl+'?OrderBy=count asc">访问统计 ↓</a>');
							}else if(sOrder=="count asc"){
								Response.Write('<a href="'+sUrl+'?OrderBy=count desc">访问统计 ↑</a>');
							}else{
								Response.Write('<a href="'+sUrl+'?OrderBy=count desc">访问统计</a>');
							}
						%>
					</td>
				</tr>
          <%
			for(var i=0;i<aClientPage.length;++i){
				var sTmpText = aClientPage[i][0].replace(/"/g,"&quot;");
				iTmpValue = aClientPage[i][1];
				var sLastTime = aClientPage[i][2];
				var iWidthPersent = Math.round(iTmpValue/iPageSum*(iWidth-5)*100)/100+1;
				var iTotalPersent = Math.round(iTmpValue/iPageSum*10000)/100;
				Response.Write(
					  '<tr title="访问次数: '+iTmpValue+' ('+iTotalPersent+'%)\n最后时间: '+sLastTime+'\n访问页: '+sTmpText.substr(0,35)+'...">'
					+ '<td class="InnerMain" style="padding-left:10px;text-align:left;"><div style="width:430px;height:15px"><nobr><a href="'+sTmpText+'" target="_blank">'+sTmpText+'</a></nobr></div></td>'
					+ '<td class="InnerMain" style="padding-left:10px;text-align:left;'+(sOrder.indexOf('time')>-1?'background-color:#eeeeee':'')+'"><div style="width:120px;height:15px"><nobr>'+sLastTime+'</nobr></div></td>'
					+ '<td width="'+iWidth+'" class="InnerMain" style="text-align:left;'+(sOrder.indexOf('count')>-1?'background-color:#eeeeee':'')+'"><div style="width:'+iWidth+'px;height:18px"><nobr>'
					+ '<span class="'+(iTmpValue==iPageMax?'HotColumnH':'ColumnH')+'" style="width:'+iWidthPersent+'px">'+(iTotalPersent>10?iTmpValue+'&nbsp;':'')+'</span>&nbsp;'
					+ '<span style="vertical-align:middle;font-size:8pt">('+iTotalPersent+'%)</span></nobr></div>'
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
	<!--#include file="../static/bott.htm"-->
</body>
</html>
