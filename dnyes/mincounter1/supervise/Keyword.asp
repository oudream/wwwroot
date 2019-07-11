<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<script language="vbscript" runat="server">
	Function AnsiCode(vstrIn)
		Dim i, strReturn, innerCode, ThisChr
		Dim Hight8, Low8
		strReturn = "" 
		For i = 1 To Len(vstrIn) 
			ThisChr = Mid(vStrIn,i,1) 
			If Abs(Asc(ThisChr)) < &HFF Then 
				strReturn = strReturn & ThisChr 
			Else
				innerCode = Asc(ThisChr)
				If innerCode < 0 Then
					innerCode = innerCode + &H10000
				End If
				Hight8 = (innerCode And &HFF00) \ &HFF
				Low8 = innerCode And &HFF
				strReturn = strReturn & "%" & Hex(Hight8) & "%" & Hex(Low8)
			End If 
		Next 
		AnsiCode = strReturn 
	End Function
	
	Function DeCodeAnsi(s)
		Dim i, sTmp, sResult, sTmp1
		sResult = ""
		For i=1 To Len(s)
			If Mid(s,i,1)="%" Then
				sTmp = "&H" & Mid(s,i+1,2)
				If isNumeric(sTmp) Then
					If CInt(sTmp)=0 Then
						i = i + 2
					ElseIf CInt(sTmp)>0 And CInt(sTmp)<128 Then
						sResult = sResult & Chr(sTmp)
						i = i + 2
					Else
						If Mid(s,i+3,1)="%" Then
							sTmp1 = "&H" & Mid(s,i+4,2)
							If isNumeric(sTmp1) Then
								sResult = sResult & Chr(CInt(sTmp)*16*16 + CInt(sTmp1))
								i = i + 5
							End If
						Else
							sResult = sResult & Chr(sTmp)
							i = i + 2
						End If
					End If
				Else
					sResult = sResult & Mid(s,i,1)
				End If
			Else
				sResult = sResult & Mid(s,i,1)
			End If
		Next
		DeCodeAnsi = sResult
	End Function
</script>
<%
	oConn.Open();
	
	var bIsLogined = chkLogin(true);
	
	var iWidth = 150;

	var aClientKeyword = new Array();
	var iKeywordSum=0, iKeywordMax=0,iTmpValue=0;
	
	var iFound = 0;
	var sOrderBy = '';
	var sSearchKeyword = ""
	if(Request.QueryString('SearchKeyword').Count>0)	sSearchKeyword = Request.QueryString('SearchKeyword')+'';
	var sSearchKeyword1 = AnsiCode(sSearchKeyword);
	if(sSearchKeyword) sOrderBy += 'instr(Keyword,"'+sSearchKeyword+'") desc,'
	if(sSearchKeyword1) sOrderBy += 'instr(Keyword,"'+sSearchKeyword1+'") desc,'
	
	var sOrder
	if(Request.QueryString('OrderBy').Count>0) sOrder = Request.QueryString('OrderBy')(1).toLowerCase();
	switch(sOrder){
	case "time desc":
		sOrderBy += "CreateTime DESC"; break;
	case "time asc":
		sOrderBy += "CreateTime ASC"; break;
	case "count desc":
		sOrderBy += "KeywordCount DESC"; break;
	case "count asc":
		sOrderBy += "KeywordCount ASC"; break;
	default:
		sOrder = "count desc";
		sOrderBy += "KeywordCount desc"; break;
	}
	
	sSqlString = "select top 20 Keyword,KeywordCount,CreateTime,RefPage from t_Keyword where SiteId=" + iSiteId + ' order by '+sOrderBy+'';
	var oRs = oConn.Execute( sSqlString );
	if(!oRs.EOF){
		var aRs = oRs.GetRows(20);
		for(var i=0;i<=aRs.ubound(2);++i){
			var sTmpText = DeCodeAnsi(aRs.getItem(0,i));
			if(sTmpText.indexOf(sSearchKeyword)>-1) iFound++;
			iTmpValue = aRs.getItem(1,i);
			var sLastTime = formatDateTime(new Date(aRs.getItem(2,i)),0);
			var sRefPage = aRs.getItem(3,i);
			iKeywordMax = (iTmpValue>iKeywordMax?iTmpValue:iKeywordMax);
			aClientKeyword[aClientKeyword.length] = Array(sTmpText,iTmpValue,sLastTime,sRefPage);
		}
		sSqlString = "select sum(KeywordCount) from t_Keyword where SiteId=" + iSiteId
		iKeywordSum = oConn.Execute(sSqlString)(0).Value;
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
<title>关键字统计 - <%=WebTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="../style/main.css" rel="stylesheet" type="text/css" />
</head>

<body>
<!--#include file="inc_head.asp"-->
	<table width="760" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="30%" align="left" valign="top"><p>
				<fieldset style="padding:3px">
					<legend>查询</legend>
					<form name="formSearchKeyword" method="get" class="OuterHead" style="height:24px">
					<table class="SubHeader"><tr><td>
					　关键字 
					  <input type="text" name="SearchKeyword" class="ThinBorderText" style="width:100px;" value="<%=sSearchKeyword%>">
						<input type="button" value="查询" onClick="form.submit();" class="button">
						<%=(sSearchKeyword?"<center><font color=red>找到 " + iFound + " 条记录。</font></center>":"")%>
					</td></tr></table>
					</form>
				</fieldset>
      <p>*说明： 所有数据皆为已知数据，未知数据不在统计之列。</p>
      <p>COCOON Counter 将从下列搜索引擎所连入的连接中分析出搜索关键字：GOOGLE、SINA、SOHU、YAHOO、3721、百度、网易、OpenFind、Lycos、AOL、ONSEEK 。 </p>
      <p>由于各搜索引擎所使用的内码不完全相同，所以可能出现关键字重复的现象。（如：google用的是utf-8，而新浪、搜狐等用的是gb2312）</p>
      <p align="right">-- Sunrise_Chen </p></td>
      <td align="right" valign="top"><table border="0" cellpadding="2" cellspacing="1" class="OuterTable" width="">
        <tr>
          <td class="OuterHead"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
              <tr>
                <td nowrap="nowrap" style="text-align:left">
				<span id="ItemTitle"><font face="webdings">8</font> 关键字统计 (前20位)</span>
				</td>
                <td align="right" class="CodingInfo">Coding by Sunrise_Chen. </td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td><table width="100%"  border="0" cellpadding="0" cellspacing="0" class="OuterTable">
              <tr>
                <td class="InnerHead">关键字</td>
                <td class="InnerHead">引用页</td>
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
			for(var i=0;i<aClientKeyword.length;++i){
				var sTmpText = aClientKeyword[i][0].replace(/"/g,"&quot;");
				iTmpValue = aClientKeyword[i][1];
				var sLastTime = aClientKeyword[i][2];
				var sRefPage = aClientKeyword[i][3].replace(/"/g,"&quot;");
				var iWidthPersent = Math.round(iTmpValue/iKeywordSum*(iWidth-5)*100)/100+1;
				var iTotalPersent = Math.round(iTmpValue/iKeywordSum*10000)/100;
				Response.Write(
					  '<tr title="关键字: '+sTmpText+'\n次数: '+iTmpValue+' ('+iTotalPersent+'%)\n访问时间: '+sLastTime+'\n引用页: '+sRefPage.substr(0,35)+'...">'
					+ '<td class="InnerMain" style="padding-left:5px;text-align:left"><div style="width:100px;height:18px"><nobr>'+sTmpText+'</nobr></div></td>'
					+ '<td class="InnerMain" style="padding-left:5px;text-align:left"><div style="width:100px;height:18px"><nobr><a href="'+sRefPage+'" target="_blank">'+sRefPage.replace('http://','').replace(/\/.*/,'')+'</a></nobr></div></td>'
					+ '<td class="InnerMain" style="padding-left:5px;text-align:left"><div style="width:118px;height:18px"><nobr>'+sLastTime+'</nobr></div></td>'
					+ '<td width="'+iWidth+'" class="InnerMain" style="text-align:left"><div style="width:'+iWidth+'px;height:18px"><nobr>'
					+ '<span class="'+(iTmpValue==iKeywordMax?'HotColumnH':'ColumnH')+'" style="width:'+iWidthPersent+'px">'+(iTotalPersent>10?iTmpValue+'&nbsp;':'')+'</span>&nbsp;'
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
      </table></td>
    </tr>
  </table>
<!--#include file="../static/bott.htm"-->
</body>
</html>
