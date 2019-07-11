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
	var sSqlString = "select Year(CountDate),Month(CountDate),Sum(HourCount),Sum(PageViewCount) from v_Count "
		+ "where siteid="+iSiteId+" and (CountDate between #"+(dToday.getYear())+"-1-1# and #"+(dToday.getYear())+"-12-31#) "
		+ "group by Year(CountDate), Month(CountDate)"
	var oRs = oConn.Execute( sSqlString );
	for(var i=0;i<iMonthOfYear;++i){
		var sTmpDate = dToday.getYear() + " 年 " + (i+1) + " 月";
		if(!oRs.EOF){
			if(oRs.fields.item(1).value==(i+1)){
				aDay[aDay.length] = Array(sTmpDate,parseInt(oRs.Fields.Item(2).Value),parseInt(oRs.Fields.Item(3).Value));
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
	var sNowFromTo = ' '+(dToday.getYear())+'-1-1 ~ '+(dToday.getYear())+'-12-31';
	
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
<title>年报一览 - <%=WebTitle%></title>
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
			
			<span id="ItemTitle"><font face="webdings">8</font> 年流量统计</span>
			<span class="SubHeader">&nbsp;
			<%="Scope: "+sNowFromTo+"  ("+iMonthOfYear+" Months), &nbsp; Total: "+iDaySum+" / "+iPageDaySum+" ."%>
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
				<span class="box" style="display:none"><span class="box2">
					<select name="selDate">
						<script language="javascript">for(var i=1;i<=12;++i) document.write("<option value="+i+">"+i+"月</option>")</script>
					</select>
				</span></span>
				<span class="box" style="display:none"><span class="box2">
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
			<% for(var j=formatMaxValue(iPageDayMax),i=j-(j/5);i>=0;i-=j/5){ %>
				<tr>
					<td height="30" align="right" valign="bottom"><span class="NumberFlag"><%=i%></span>&nbsp;</td>
				</tr>
			<% } %>
			</table>
		</td>
        <td class="InnerMain" title="[年访问情况一览]<%='\n'+sNowFromTo+'\n总访问量: '+iDaySum+' 次\n总浏览量: '+iPageDaySum+' 次'%>">
		<table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr class="styTdBGLine">
					<%
						for(var i=0;i<aDay.length;++i){
							//var sTmpDate = formatDateTime(aDay[i][0],2);
							var iPageHeightPersent = Math.round(150*(aDay[i][2]/formatMaxValue(iPageDayMax))*100)/100+1;
							var iHeightPersent = Math.round(150*(aDay[i][1]/formatMaxValue(iPageDayMax))*100)/100+1;
							iPageHeightPersent = iPageHeightPersent - iHeightPersent;
							if(iPageHeightPersent<=0) iPageHeightPersent=1;
							Response.Write('<td class="TimeColumn">'
								+ '<span title="时段　 : '+(sTmpDate)+' \n'
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
					Response.Write("<td align=center style='padding-top:3px'><a href='Day.asp?QueryDate="+(formatDateTime(new Date(dToday.getYear(),i+1,0),2))+"'>"+getMonthName(i)+"</a></td>\n");
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
						 title="上一年" 
						 onFocus="blur()" 
						 onClick="location.href='<%=sUrl+'?QueryDate='+formatDateTime(new Date(dToday.getYear()-1,dToday.getMonth(),1),2)%>'"
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
						 title="下一年" 
						 onFocus="blur()" 
						 onClick="location.href='<%=sUrl+'?QueryDate='+formatDateTime(new Date(dToday.getYear()+1,dToday.getMonth(),1),2)%>'"
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

<!--#include file="../static/bott.htm"-->
</body>
</html>
