<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<%
	oConn.Open();
	
	var bIsLogined = chkLogin(true);
	
	var iDaySum = 0;
	var iPageDaySum = 0;
	var iDayMax = 0;
	var iPageDayMax = 0;
	var aDay = new Array();
	var aPageNowHour = new Array();
	var sSqlString = "select CountDate,HourCount,PageViewCount from v_Count "
		+ "where siteid="+iSiteId+" and (CountDate between #"+sLastWeek+"# and #"+sToday+"#) order by CountDate asc";
	var oRs = oConn.Execute( sSqlString );
	for(var i=0;i<iDayOfWeek;++i){
		var sTmpDate = new Date((1000*60*60*24) * i + dLastWeek.valueOf());
		if(!oRs.EOF){
			var sTmp = formatDateTime(new Date(oRs.Fields.Item(0).Value),2);
			if(formatDateTime(sTmpDate,2)==sTmp){
				aDay[aDay.length] = Array(sTmpDate,parseInt(oRs.Fields.Item(1).Value),parseInt(oRs.Fields.Item(2).Value));
				oRs.moveNext();
			}else{
				aDay[aDay.length] = Array(sTmpDate,0,0);
			}
		}else{
			aDay[aDay.length] = Array(sTmpDate,0,0);
		}
		if(iDayMax<aDay[aDay.length-1][1]) iDayMax=aDay[aDay.length-1][1];
		if(iPageDayMax<aDay[aDay.length-1][2]) iPageDayMax=aDay[aDay.length-1][2];
		iDaySum += aDay[aDay.length-1][1]; iPageDaySum += aDay[aDay.length-1][2];
	}
	var sNowFromTo = ' '+(sLastWeek)+' ~ '+(sToday)+'';
	
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
	var sSqlString = "SELECT WeekDay(CountDate),Sum(HourCount),Sum(PageViewCount) FROM v_Count"
		+ " WHERE siteid="+iSiteId
		+ " GROUP BY WeekDay(CountDate)"
		+ " ORDER BY WeekDay(CountDate) ASC";
	var oRs = oConn.Execute( sSqlString );
	for(var i=0;i<iDayOfWeek;++i){
		if(!oRs.EOF){
			if(oRs.Fields.Item(0).Value==i+1){
				aTotalHour[aTotalHour.length] = Array(i+1,oRs.Fields.Item(1).Value,oRs.Fields.Item(2).Value);
				oRs.MoveNext()
			}else{
				aTotalHour[aTotalHour.length] = Array(i+1,0,0);
			}
		}else{
			aTotalHour[aTotalHour.length] = Array(i+1,0,0);
		}
		iSumTotalHour += aTotalHour[aTotalHour.length-1][1];
		iPageSumTotalHour += aTotalHour[aTotalHour.length-1][2];
		if(iTotalHourMax<aTotalHour[aTotalHour.length-1][1]) iTotalHourMax=aTotalHour[aTotalHour.length-1][1];
		if(iPageTotalHourMax<aTotalHour[aTotalHour.length-1][2]) iPageTotalHourMax=aTotalHour[aTotalHour.length-1][2];
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
<title>周报一览 - <%=WebTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="../style/main.css" rel="stylesheet" type="text/css" />
</head>

<body>
<!--#include file="inc_head.asp"-->
<table width="760" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="50%" align="left"><table width="375" border="0" cellpadding="2" cellspacing="1" class="OuterTable">
      <tr>
        <td class="OuterHead"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
            <tr>
              <td nowrap="nowrap" style="text-align:left"> <span id="ItemTitle"><font face="webdings">8</font> 一周流量统计</span> <span class="SubHeader">&nbsp; <%="Scope: "+sNowFromTo+" ."%> </span> </td>
            </tr>
        </table></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellpadding="0" cellspacing="1" class="InnerTable">
            <tr>
              <td width="70" class="InnerHead">
                <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <% for(var j=formatMaxValue(iPageDayMax),i=j-(j/5);i>=0;i-=j/5){ %>
                  <tr>
                    <td height="30" align="right" valign="bottom"><span class="NumberFlag"><%=i%></span>&nbsp;</td>
                  </tr>
                  <% } %>
              </table></td>
              <td class="InnerMain" title="[一周情况一览]<%='\n'+sNowFromTo+'\n总访问量: '+iDaySum+' 次\n总浏览量: '+iPageDaySum+' 次'%>">
                <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr class="styTdBGLine">
                    <%
						for(var i=0;i<aDay.length;++i){
							var sTmpDate = formatDateTime(aDay[i][0],2);
							var iPageHeightPersent = Math.round(150*(aDay[i][2]/formatMaxValue(iPageDayMax))*100)/100+1;
							var iHeightPersent = Math.round(150*(aDay[i][1]/formatMaxValue(iPageDayMax))*100)/100+1;
							iPageHeightPersent = iPageHeightPersent - iHeightPersent;
							if(iPageHeightPersent<=0) iPageHeightPersent=1;
							Response.Write('<td class="TimeColumn" bgColor='+(aDay[aDay.length-1][0].getDate()==iDayOfWeek?(aDay[i][0].getDate()<11||aDay[i][0].getDate()>20?'fafafa':'white'):'white')+'>'
								+ '<span title="时段　 : '+(sTmpDate)+' (星期'+aDay[i][0].getCnWeekName()+')\n'
								+ '浏览量 : ' + aDay[i][2]+' 次 ('+Math.round((aDay[i][2]/iPageDaySum*100)*100)/100+'%)\n'
								+ '访问量 : ' + aDay[i][1]+' 次 ('+Math.round((aDay[i][1]/iDaySum*100)*100)/100+'%)">'
								+ '<div class="Column1" style="height:'+iPageHeightPersent+'px;">'
								+ '<span class="TimeWord">&nbsp;'+(iPageHeightPersent>20?aDay[i][2]:'')+'</span></div>'
								+ '<div class="Column" style="height:'+iHeightPersent+'px;">'
								+ '<span class="TimeWord">&nbsp;'+(iHeightPersent>20?aDay[i][1]:'')+'</span></div>');
							Response.write('</span></td>\n');
						}
					%>
                  </tr>
                  <tr align="center">
                    <% 
				for(var i=0;i<aDay.length;++i){
					if(aDay[i][0].getDay()<1||aDay[i][0].getDay()>5)
						Response.Write("<td align=center bgColor='#eeeeee' style='padding-top:3px'><a href='Hour.asp?QueryDate="+formatDateTime(aDay[i][0],2)+"'>"+(aDay[i][0].getCnWeekName())+"</a></td>\n");
					else
						Response.Write("<td align=center style='padding-top:3px'><a href='Hour.asp?QueryDate="+formatDateTime(aDay[i][0],2)+"'>"+(aDay[i][0].getCnWeekName())+"</a></td>\n");
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
						 title="上一周" 
						 onFocus="blur()" 
						 onClick="location.href='<%=sUrl+'?QueryDate='+formatDateTime(new Date(dToday - (1000*60*60*24)*7),2)%>'"
						 
						>
                    </td>
                  </tr>
                  <tr class="MainText">
                    <td align="center" valign="middle"> </td>
                  </tr>
                  <tr class="MainText">
                    <td height="26" align="center" valign="bottom">
                      <input
						 name="Submit" 
						 type="button" 
						 class="stySmallBtnBott" 
						 id="Submit6" 
						 value="6" 
						 title="下一周" 
						 onFocus="blur()" 
						 onClick="location.href='<%=sUrl+'?QueryDate='+formatDateTime(new Date(dToday.valueOf() + (1000*60*60*24)*7),2)%>'"
						 
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
    </table></td>
    <td align="right"><table width="375" border="0" cellpadding="2" cellspacing="1" class="OuterTable">
      <tr>
        <td class="OuterHead"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
            <tr>
              <td nowrap="nowrap" style="text-align:left"> <span id="ItemTitle"><font face="webdings">8</font> 所有一周访问情况一览</span> <span class="SubHeader">&nbsp; <%=""+"Total: "+iSumTotalHour+"/"+iPageSumTotalHour+"."%> </span> </td>
              <td align="right" class="CodingInfo"></td>
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
              <td class="InnerMain" title="[所有一周访问情况一览]<%='\n\n总访问量: '+iSumTotalHour+' 次\n总浏览量: '+iPageSumTotalHour+' 次'%>">
                <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr class="styTdBGLine">
                    <%
						for(var i=0;i<iDayOfWeek;++i){
							var iPageHeightPersent = Math.round(150*(aTotalHour[i][2]/formatMaxValue(iPageTotalHourMax?iPageTotalHourMax:iTotalHourMax))*100)/100+1;
							var iHeightPersent = Math.round(150*(aTotalHour[i][1]/formatMaxValue(iPageTotalHourMax?iPageTotalHourMax:iTotalHourMax))*100)/100+1;
							iPageHeightPersent = iPageHeightPersent - iHeightPersent;
							if(iPageHeightPersent<=0) iPageHeightPersent=1
							Response.Write('<td class="TimeColumn" bgColor='+(i<10||i>19?'fafafa':'white')+'>'
								+ '<span title="时间段: '+(i+1)+' 日\n'
								+ '浏览量: '+aTotalHour[i][2]+' ('+Math.round((aTotalHour[i][2]/iPageSumTotalHour*100)*100)/100+'%)\n'
								+ '访问量: '+aTotalHour[i][1]+' ('+Math.round((aTotalHour[i][1]/iSumTotalHour*100)*100)/100+'%)">'
								+ '<div class="Column1" style="height:'+iPageHeightPersent+'px;">'
								+ '<span class="TimeWord">&nbsp;'+(iPageHeightPersent>30?aTotalHour[i][2]:'')+'</span></div>'
								+ '<div class="Column" style="height:'+iHeightPersent+'px;">'
								+ '<span class="TimeWord">&nbsp;'+(iHeightPersent>30?aTotalHour[i][1]:'')+'</span></div>'
								+ '</span></td>\n'
							);
						}
					%>
                  </tr>
                  <tr align="center">
                    <% 
						for(var i=0;i<iDayOfWeek;++i){
							Response.Write("<td style='padding-top:3px'>"+(Array("日","一","二","三","四","五","六")[i])+"</td>\n");
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
    </table></td>
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

<!--#include file="../static/bott.htm"-->
</body>
</html>
