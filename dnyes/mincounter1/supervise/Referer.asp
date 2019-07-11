<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<%
	oConn.Open();
	
	var bIsLogined = chkLogin(true);
	
	var iWidth = 170;

	var aClientRefSite = new Array();
	var iRefSiteSum=0, iRefSiteMax=0,iTmpValue=0;
	
	var iFound = 0;
	var sOrder = '';
	var sOrderBy = '';
	var sSearchRefSite = '';
	if(Request.Form('SearchRefSite').Count>0)	sSearchRefSite = Request.Form('SearchRefSite')(1);
	if(Request.QueryString('OrderBy').Count>0) sOrder = Request.QueryString('OrderBy')(1).toLowerCase();
	switch(sOrder){
	case "time desc":
		sOrderBy = "CreateTime DESC"; break;
	case "time asc":
		sOrderBy = "CreateTime ASC"; break;
	case "count desc":
		sOrderBy = "RefPageCount DESC"; break;
	case "count asc":
		sOrderBy = "RefPageCount ASC"; break;
	default:
		sOrder = "count desc";
		sOrderBy = "RefPageCount DESC"; break;
	}
	if(sSearchRefSite) sOrderBy = 'instr(RefSite,"'+sSearchRefSite+'") desc,RefPageCount DESC'	
	
	sSqlString = "select top 20 RefSite,RefPageCount,CreateTime,RefPage,id from t_RefPage where SiteId=" + iSiteId + ' order by '+sOrderBy;
	var oRs = oConn.Execute( sSqlString );
	if(!oRs.EOF){
		var aRs = oRs.GetRows(20);
		for(var i=0;i<=aRs.ubound(2);++i){
			var sTmpText = aRs.getItem(0,i).replace('http://','');
			if(sTmpText.indexOf(sSearchRefSite)>-1) iFound++;
			iTmpValue = aRs.getItem(1,i);
			var sLastTime = formatDateTime(new Date(aRs.getItem(2,i)),0);
			var sRefPage = aRs.getItem(3,i);
			iRefSiteMax = (iTmpValue>iRefSiteMax?iTmpValue:iRefSiteMax);
			aClientRefSite[aClientRefSite.length] = Array(sTmpText,iTmpValue,sLastTime,sRefPage);
		}
			
		sSqlString = "select sum(RefPageCount) from t_RefPage where SiteId=" + iSiteId;
		iRefSiteSum = oConn.Execute(sSqlString)(0).Value;
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
<title>来源统计 - <%=WebTitle%></title>
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
				<span id="ItemTitle"><font face="webdings">8</font> 来源统计 (前20位)</span></td>
            <td nowrap="nowrap" style="text-align:left;width:400px">
				<form name="formSearchRefSite" method="post" class="SubHeader">
              　关键字
                <input type="text" name="SearchRefSite" class="ThinBorderText" style="width:130px" value="<%=sSearchRefSite%>">
                <input name="button" type="button" onClick="form.submit();" value="查询" class="button">
								&nbsp;&nbsp;&nbsp;<font color="red"><%=(sSearchRefSite?"找到 " + iFound + " 条记录。":"")%></font>
              </form>						</td>
            <td align="right" class="CodingInfo">Coding by Sunrise_Chen. </td>
          </tr>
      </table></td>
    </tr>
    <tr>
      <td><table width="100%"  border="0" cellpadding="0" cellspacing="0" class="OuterTable">
				<tr>
					<td class="InnerHead">引用站站点</td>
					<td class="InnerHead">引用连接</td>
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
			for(var i=0;i<aClientRefSite.length;++i){
				var sTmpText = aClientRefSite[i][0].replace(/"/g,"&quot;");
				if(!sTmpText.length) sTmpText="输入地址或从书签进入"
				iTmpValue = aClientRefSite[i][1];
				var sLastTime = aClientRefSite[i][2];
				var sRefPage = aClientRefSite[i][3].replace(/"/g,"&quot;");
				var iWidthPersent = Math.round(iTmpValue/iRefSiteSum*(iWidth-5)*100)/100+1;
				var iTotalPersent = Math.round(iTmpValue/iRefSiteSum*10000)/100;
				Response.Write(
					  '<tr title="引用站点: '+sTmpText+'\n访问次数: '+iTmpValue+' ('+iTotalPersent+'%)\n最后时间: '+sLastTime+'\n引用连接: '+sRefPage.substr(0,35)+'...">'
					+ '<td class="InnerMain" style="padding-left:10px;text-align:left"><div style="width:130px;height:15px"><nobr>'+(sTmpText.indexOf('.')>0?'<a href="http://'+sTmpText+'" target="_blank">'+sTmpText+'</a>':sTmpText)+'</nobr></div></td>'
					+ '<td class="InnerMain" style="padding-left:10px;text-align:left;"><div style="width:295px;height:15px"><nobr><a href="'+sRefPage+'" target="_blank">'+sRefPage+'</a></nobr></div></td>'
					+ '<td class="InnerMain" style="padding-left:10px;text-align:left;'+(sOrder.indexOf('time')>-1?'background-color:#eeeeee':'')+'"><div style="width:120px;height:15px"><nobr>'+sLastTime+'</nobr></div></td>'
					+ '<td width="'+iWidth+'" class="InnerMain" style="text-align:left;'+(sOrder.indexOf('count')>-1?'background-color:#eeeeee':'')+'"><div style="width:'+iWidth+'px;height:18px"><nobr>'
					+ '<span class="'+(iTmpValue==iRefSiteMax?'HotColumnH':'ColumnH')+'" style="width:'+iWidthPersent+'px">'+(iTotalPersent>10?iTmpValue+'&nbsp;':'')+'</span>&nbsp;'
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
