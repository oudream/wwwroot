<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<%
	oConn.Open();
	
	var bIsLogined = chkLogin(true);
	
	if(oQueryDate.Count==1){
		dYesterday = dToday;
		dToday = new Date(dToday.valueOf() + (1000*60*60*24));
		sToday = formatDateTime(dToday,2);
		sYesterday = formatDateTime(dYesterday,2);
		iNowHour = 0;
	}
	
	var iSumNowHour = 0;
	var iPageSumNowHour = 0;
	var aNowHour = new Array();
	var aPageNowHour = new Array();
	var sSqlString = "select h_0, h_1, h_2, h_3, h_4, h_5, h_6, h_7, h_8, h_9, h_10, h_11, h_12, "
		+ "h_13, h_14, h_15, h_16, h_17, h_18, h_19, h_20, h_21, h_22, h_23, CountDate, "
		+ "s_0, s_1, s_2, s_3, s_4, s_5, s_6, s_7, s_8, s_9, s_10, s_11, s_12, "
		+ "s_13, s_14, s_15, s_16, s_17, s_18, s_19, s_20, s_21, s_22, s_23 "
		+ "from t_Count where siteid="+iSiteId+" and (CountDate=#"+sYesterday+"# or CountDate=#"+sToday+"#) order by id asc";
	var oRs = oConn.Execute( sSqlString );
	if(!oRs.EOF){
		if(formatDateTime(new Date(oRs.Fields.Item(24).Value),2)==sYesterday){
			while(!oRs.EOF){
				for(var i=0;i<24;++i){
					aNowHour[aNowHour.length] = oRs.fields.item(i).value;
					aPageNowHour[aPageNowHour.length] = oRs.fields.item(i+25).value;
				}
				oRs.MoveNext();
			}
		}else{
			for(var i=0;i<24;++i){
				aNowHour[aNowHour.length] = 0;
				aPageNowHour[aPageNowHour.length] = 0;
			}
			for(i=0;i<24;++i){
				aNowHour[aNowHour.length] = oRs.fields.item(i).value;
				aPageNowHour[aPageNowHour.length] = oRs.fields.item(i+25).value;
			}
		}
	}
	for(var i=aNowHour.length;i<48;++i){
		aNowHour[aNowHour.length]=0;
		aPageNowHour[aPageNowHour.length]=0;
	}
	aNowHour = aNowHour.slice(iNowHour+1,iNowHour+1+24);
	aPageNowHour = aPageNowHour.slice(iNowHour+1,iNowHour+1+24);
	var iNowHourMax = eval("Math.max("+aNowHour.toString()+")");
	var iPageNowHourMax = eval("Math.max("+aPageNowHour.toString()+")");
	for(var i=0;i<aNowHour.length;++i){
		iSumNowHour+=aNowHour[i];
		iPageSumNowHour+=aPageNowHour[i];
	}
	var sNowFromTo = sYesterday+' '+(iNowHour)+':00 ~ '+sToday+' '+(iNowHour)+':00';
	
	var iCountDate = 0;
	var iSumTotalHour = 0;
	var iPageSumTotalHour = 0;
	var dCountFromDate = formatDateTime(new Date(),2);
	var dCountToDate = formatDateTime(new Date(),2);
	var aTotalHour = new Array();
	var aPageTotalHour = new Array();
	var iTotalHourMax = 0;
	var iPageTotalHourMax = 0;
	var sTotalFromTo = '';
	var sSqlString = "SELECT Sum(h_0),Sum(h_1),Sum(h_2),Sum(h_3),Sum(h_4),Sum(h_5),"
		+ "Sum(h_6),Sum(h_7),Sum(h_8),Sum(h_9),Sum(h_10),Sum(h_11),"
		+ "Sum(h_12),Sum(h_13),Sum(h_14),Sum(h_15),Sum(h_16),Sum(h_17),"
		+ "Sum(h_18),Sum(h_19),Sum(h_20),Sum(h_21),Sum(h_22),Sum(h_23),"
		+ "Max(CountDate)-Min(CountDate),Min(CountDate),Max(CountDate), "
		+ "Sum(s_0),Sum(s_1),Sum(s_2),Sum(s_3),Sum(s_4),Sum(s_5),"
		+ "Sum(s_6),Sum(s_7),Sum(s_8),Sum(s_9),Sum(s_10),Sum(s_11),"
		+ "Sum(s_12),Sum(s_13),Sum(s_14),Sum(s_15),Sum(s_16),Sum(s_17),"
		+ "Sum(s_18),Sum(s_19),Sum(s_20),Sum(s_21),Sum(s_22),Sum(s_23)"
		+ "FROM t_Count WHERE siteid=" + iSiteId + " GROUP BY siteid";
	var oRs = oConn.Execute( sSqlString );
	if(!oRs.EOF){
		for(var i=0;i<24;++i){
			aTotalHour[i] = oRs.Fields.Item(i).Value;
			aPageTotalHour[i] = oRs.Fields.Item(i+27).Value;
			iSumTotalHour += aTotalHour[i];
			iPageSumTotalHour += aPageTotalHour[i];
		}
		iCountDate = oRs.Fields.Item(24).Value;
		dCountFromDate = formatDateTime(new Date(oRs.Fields.Item(25).Value),2);
		dCountToDate = formatDateTime(new Date(oRs.Fields.Item(26).Value),2);
		iTotalHourMax = eval("Math.max("+aTotalHour.toString()+")");
		iPageTotalHourMax = eval("Math.max("+aPageTotalHour.toString()+")");
		sTotalFromTo = dCountFromDate+' ~ '+dCountToDate
	}else{
		for(var i=0;i<24;++i){
			aTotalHour[i] = 0;
			aPageTotalHour[i] = 0;
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
<title>日报一览 - <%=WebTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="../style/main.css" rel="stylesheet" type="text/css" />
</head>

<body>
<!--#include file="inc_head.asp"-->
<table width="760" border="0" cellpadding="2" cellspacing="1" class="OuterTable">
  <tr>
    <td class="OuterHead"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
      <tr>
        <td nowrap="nowrap" style="text-align:left">
			<span id="ItemTitle"><font face="webdings">8</font> 24小时流量统计</span>
			<span class="SubHeader">&nbsp;
			<%="Scope: "+sNowFromTo+", &nbsp; "+"Total: "+iSumNowHour+"/"+iPageSumNowHour+"."%>
			</span>
		</td>
        <td align="right" nowrap>
					<form name="formQuery" id="formQuery" method="get" style="display:none">
						<script language="javascript">
							function doQuery(oForm){
								with(oForm){
									QueryDate.value = selDate[0].value + '-' + selDate[1].value + '-' + selDate[2].value;
									for(var i=0;i<elements.length;++i) elements[i].disabled = true;
									QueryDate.disabled = false;
									submit();
								}
							}
						</script>
						<span class="box"><span class="box2">
							<select name="selDate">
								<script language="javascript">for(var i=(new Date()).getFullYear();i>=2002;--i) document.write("<option value="+i+">"+i+"年</option>")</script>
							</select>
						</span></span>
						<span class="box"><span class="box2">
							<select name="selDate">
								<script language="javascript">for(var i=1;i<=12;++i) document.write("<option value="+i+">"+i+"月</option>")</script>
							</select>
						</span></span>
						<span class="box"><span class="box2">
							<select name="selDate">
								<script language="javascript">for(var i=1;i<=31;++i) document.write("<option value="+i+">"+i+"日</option>")</script>
							</select>
						</span></span>
						<input type="hidden" name="QueryDate" value="">
						<input type="button" value="GO" class="button" onFocus="blur()" onClick="doQuery(form)">
					</form>
					<input type="button" class="button" onFocus="blur()" value="历史查询" style="width:75" onClick="document.getElementById('formQuery').style.display='';style.display='none'">
				</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellpadding="0" cellspacing="1" class="InnerTable">
      <tr>
        <td width="70" class="InnerHead">
			<table width="100%"  border="0" cellspacing="0" cellpadding="0">
			<% for(var j=formatMaxValue(iPageNowHourMax),i=j-(j/5);i>=0;i-=j/5){ %>
				<tr>
					<td height="30" align="right" valign="bottom"><span class="NumberFlag"><%=i%></span>&nbsp;</td>
				</tr>
			<% } %>
			</table>
		</td>
        <td class="InnerMain" title="[最近24小时访问情况]<%='\n'+sNowFromTo+'\n总访问量: '+iSumNowHour+' 次\n总浏览量: '+iPageSumNowHour+' 次'%>">
		<table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr class="styTdBGLine">
					<%
						for(var i=0;i<24;++i){
							var iPageHeightPersent = Math.round(150*(aPageNowHour[i]/formatMaxValue(iPageNowHourMax))*100)/100;
							var iHeightPersent = Math.round(150*(aNowHour[i]/formatMaxValue(iPageNowHourMax))*100)/100;
							iPageHeightPersent = iPageHeightPersent - iHeightPersent;
							if(iPageHeightPersent<1) iPageHeightPersent=1;
							if(iHeightPersent<1) iHeightPersent=1;
							Response.Write('<td class="TimeColumn">'
								+ '<span title="时段　: [ '+(i+iNowHour+1)%24+':00 ~ '+(i+iNowHour+2)%24+':00 ]\n'
								+ '浏览量: ' + aPageNowHour[i]+' 次 ('+Math.round((aPageNowHour[i]/iPageSumNowHour*100)*100)/100+'%)\n'
								+ '访问量: ' + aNowHour[i]+' 次 ('+Math.round((aNowHour[i]/iSumNowHour*100)*100)/100+'%)">'
								+ '<div class="Column1" style="height:'+iPageHeightPersent+'px;">'
								+ '<span class="TimeWord">&nbsp;'+(iPageHeightPersent>20?aPageNowHour[i]:'')+'</span></div>'
								+ '<div class="Column" style="height:'+iHeightPersent+'px;">'
								+ '<span class="TimeWord">&nbsp;'+(iHeightPersent>20?aNowHour[i]:'')+'</span></div>');
							Response.write('</span></td>\n');
						}
					%>
          </tr>
          <tr align="center">
					<% 
						for(var i=0;i<24;++i){
							Response.Write("<td align=center>"+(i+iNowHour+1)%24+"</td>\n");
						}
					%>
          </tr>
        </table></td>
				<td width="16" class="InnerMain"><table width="100%" height="165" border="0" cellpadding="0" cellspacing="0">
          <tr class="MainText">
            <td height="26" align="center" valign="top">
						<input
						 name="Submit" 
						 type="button" 
						 class="stySmallBtnTop" 
						 id="Submit5" 
						 value="5" 
						 title="前一天" 
						 onFocus="blur()" 
						 onClick="location.href='<%=sUrl+'?QueryDate='+formatDateTime(new Date(dYesterday - (1000*60*60*24)),2)%>'"
						>
						</td>
          </tr>
          <tr class="MainText">
            <td align="center" valign="middle">
						</td>
          </tr>
          <tr class="MainText">
            <td height="26" align="center" valign="bottom">
						<input
						 name="Submit" 
						 type="button" 
						 class="stySmallBtnBott" 
						 id="Submit6" 
						 value="6" 
						 title="后一天" 
						 onFocus="blur()" 
						 onClick="location.href='<%=sUrl+'?QueryDate='+formatDateTime(new Date(dToday),2)%>'"
						>
						</td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td class="OuterFoot"></td>
  </tr>
</table>

<table width="760" border="0" cellpadding="2" cellspacing="1">
  <tr>
    <td align=right>
		<img src="../images/blank.gif" class="Column1" width=20 height=10 align=absmiddle style="border:1px solid black" /> - 浏览量
		<img src="../images/blank.gif" class="Column" width=20 height=10 align=absmiddle style="border:1px solid black"/> - 访问量
	</td>
  </tr>
</table>

<table width="760" border="0" cellpadding="2" cellspacing="1" class="OuterTable">
  <tr>
    <td class="OuterHead"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
      <tr>
        <td nowrap="nowrap" style="text-align:left">
			<span id="ItemTitle"><font face="webdings">8</font> 所有24小时统计信息</span>
			<span class="SubHeader">&nbsp;
			<%="Scope: ("+sTotalFromTo+") &nbsp; "+"Total: "+iSumTotalHour+"/"+iPageSumTotalHour+"."%>
			</span>
		</td>
        <td align="right" class="CodingInfo">Coding by Sunrise_Chen. </td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%"  border="0" cellpadding="0" cellspacing="1" class="InnerTable">
      <tr>
		<td width="80" class="InnerHead">
					<table width="100%"  border="0" cellspacing="0" cellpadding="0">
            <% for(var j=formatMaxValue(iPageTotalHourMax),i=j-(j/5);i>=0;i-=j/5){ %>
            <tr>
              <td height="30" align="right" valign="bottom"><span class="NumberFlag"><%=i%></span>&nbsp;</td>
            </tr>
            <% } %>
          </table></td>
        <td class="InnerMain" title="[所有24小时访问情况]<%='\n'+sTotalFromTo+'\n总访问量: '+iSumTotalHour+' 次\n总浏览量: '+iPageSumTotalHour+' 次'%>">
		<table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr class="styTdBGLine">
					<%
						for(var i=0;i<24;++i){
							var iPageHeightPersent = Math.round(150*(aPageTotalHour[i]/formatMaxValue(iPageTotalHourMax?iPageTotalHourMax:iTotalHourMax))*100)/100;
							var iHeightPersent = Math.round(150*(aTotalHour[i]/formatMaxValue(iPageTotalHourMax?iPageTotalHourMax:iTotalHourMax))*100)/100;
							iPageHeightPersent = iPageHeightPersent - iHeightPersent;
							if(iPageHeightPersent<=1) iPageHeightPersent=1;
							if(iHeightPersent<=1) iHeightPersent=1;
							Response.Write('<td class="TimeColumn">'
								+ '<span title="时间段: [ '+i+':00 ~ '+((i+1)%24)+':00 ]\n'
								+ '浏览量: '+aPageTotalHour[i]+' ('+Math.round((aPageTotalHour[i]/iPageSumTotalHour*100)*100)/100+'%)\n'
								+ '访问量: '+aTotalHour[i]+' ('+Math.round((aTotalHour[i]/iSumTotalHour*100)*100)/100+'%)">'
								+ '<div class="Column1" style="height:'+iPageHeightPersent+'px;">'
								+ '<span class="TimeWord">&nbsp;'+(iPageHeightPersent>30?aPageTotalHour[i]:'')+'</span></div>'
								+ '<div class="Column" style="height:'+iHeightPersent+'px;">'
								+ '<span class="TimeWord">&nbsp;'+(iHeightPersent>30?aTotalHour[i]:'')+'</span></div>'
								+ '</span></td>\n'
							);
						}
					%>
          </tr>
          <tr align="center">
					<% 
						for(var i=0;i<24;++i){
							Response.Write("<td>"+i+"</td>\n");
						}
					%>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td class="OuterFoot"></td>
  </tr>
</table>
<!--#include file="../static/bott.htm"-->
</body>
</html>
