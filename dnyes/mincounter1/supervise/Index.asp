<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<!--#include file="inc_Language.asp"-->
<!--#include file="inc_TimeZone.asp"-->
<%
	oConn.Open();
	
	var bIsLogined = chkLogin(true);
	
	sSqlString = "select * from v_SiteCount where SiteId=" + iSiteId;
	var oRs = oConn.Execute( sSqlString );
	if(oRs.EOF){
		oConn.Close();
		Response.End();
	}
	var aCount = oRs.GetRows(1);
	
	var aVisit;
	sSqlString = "select CreateTime,PageUrl,RefUrl,ip,region,linktype,ClientInfo "
		+ "from t_visit where SiteId=" + iSiteId + " order by id desc";
	var oRs = new ActiveXObject("ADODB.Recordset");
	with(oRs){
		PageSize = DefaultPageSize;
		open(sSqlString,oConn,1,3);
		var iPageCount = PageCount;
		var iRecCount = RecordCount;
		var iPage = parseInt(Request.QueryString("p")+'');
		if(isNaN(iPage)||iPage<1) iPage=1;
		if(iPage>iPageCount) iPage=iPageCount;
		if(!EOF){
			AbsolutePage = iPage;
			aVisit = GetRows(20);
		}
		close();
	}
	//var oRs = oConn.Execute( sSqlString );
	
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
<title>�ۺ�ͳ�� - <%=WebTitle%></title>
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
			<span id="ItemTitle">
			<font face="webdings">8</font> վ��ſ� --  <%=aCount.getItem(20,0).toLowerCase().replace("http://","")%>
			</span>
			</td>
            <td nowrap="nowrap" style="text-align:left">&nbsp;</td>
            <td width="180" align="right" class="CodingInfo">Coding by Sunrise_Chen. </td>
          </tr>
      </table></td>
    </tr>
    <tr>
      <td><table width="100%"  border="0" cellpadding="0" cellspacing="1" class="InnerTable">
				<tr>
					<td class="InnerHead" width="15%">����ͳ��</td>
				  <td class="InnerMain" width="35%" style="padding-left:5pt;">
						����: <%=aCount.getItem(3,0)%> ��.������: <%=aCount.getItem(4,0)%> ��.</td>
					<td class="InnerHead" width="15%">վ������</td>
					<td class="InnerMain" width="35%" style="padding-left:5pt;"><%=aCount.getItem(18,0)%>
						
					</td>
				</tr>
				<tr>
				  <td class="InnerHead">��ʼ����</td>
				  <td class="InnerMain" style="padding-left:5pt;"><%=aCount.getItem(0,0)%></td>
				  <td class="InnerHead">վ��Url1</td>
				  <td class="InnerMain" style="padding-left:5pt;"><a href="<%=aCount.getItem(20,0)%>" target="_blank"><%=aCount.getItem(20,0)%></a></td>
			  </tr>
				<tr>
				  <td class="InnerHead">ͳ������</td>
				  <td class="InnerMain" style="padding-left:5pt;"><%=aCount.getItem(2,0)%> ��.</td>
				  <td class="InnerHead">վ��Url2</td>
				  <td class="InnerMain" style="padding-left:5pt;"><%=aCount.getItem(21,0)%></td>
			  </tr>
				<tr>
				  <td class="InnerHead">ƽ��ͳ��</td>
				  <td class="InnerMain" style="padding-left:5pt;">
						����: <%=Math.round(aCount.getItem(3,0)/aCount.getItem(2,0)*100)/100%> ��/��.
						&nbsp;&nbsp;
						���: <%=Math.round(aCount.getItem(4,0)/aCount.getItem(2,0)*100)/100%> ��/��.				  </td>
				  <td class="InnerHead">վ��Url3</td>
				  <td class="InnerMain" style="padding-left:5pt;"><%=aCount.getItem(22,0)%></td>
			  </tr>
				<tr>
				  <td class="InnerHead">�������</td>
				  <td class="InnerMain" style="padding-left:5pt;"><%=aCount.getItem(6,0)%>&nbsp;&nbsp;<%=aCount.getItem(5,0)%>��.</td>
				  <td class="InnerHead">վ���ǳ�</td>
				  <td class="InnerMain" style="padding-left:5pt;"><%=aCount.getItem(23,0)%></td>
			  </tr>
				<tr>
				  <td class="InnerHead">��������</td>
				  <td class="InnerMain" style="padding-left:5pt;"> ����: <%=(aCount.getItem(7,0)?aCount.getItem(7,0):0)%> ��.
				  &nbsp;&nbsp; ���: <%=(aCount.getItem(11,0)?aCount.getItem(11,0):0)%> ��.  </td>
				  <td class="InnerHead">վ������</td>
				  <td class="InnerMain" style="padding-left:5pt;"><a href="mailto:<%=aCount.getItem(24,0)%>"> <%=aCount.getItem(24,0)%></a></td>
			  </tr>
				<tr>
				  <td class="InnerHead">��������</td>
				  <td class="InnerMain" style="padding-left:5pt;"> ����: <%=(aCount.getItem(8,0)?aCount.getItem(8,0):0)%> ��.
				  &nbsp;&nbsp; ���: <%=(aCount.getItem(12,0)?aCount.getItem(12,0):0)%> ��. </td>
				  <td class="InnerHead">ͳ���ʺ�</td>
				  <td class="InnerMain" style="padding-left:5pt;"><%=aCount.getItem(16,0)%></td>
			  </tr>
				<tr>
				  <td class="InnerHead">��������</td>
				  <td class="InnerMain" style="padding-left:5pt;"> ����: <%=(aCount.getItem(9,0)?aCount.getItem(9,0):0)%> ��.
				  &nbsp;&nbsp; ���: <%=(aCount.getItem(13,0)?aCount.getItem(13,0):0)%> ��. </td>
				  <td class="InnerHead">վ����</td>
				  <td rowspan="2" valign="top" class="InnerMain" style="padding-left:5pt;padding-top:3pt;">
				  <div style="overflow:auto;height:38">
				  <%=aCount.getItem(19,0)%>
				  </div>
				  </td>
			  </tr>
				<tr>
				  <td class="InnerHead">��������</td>
				  <td class="InnerMain" style="padding-left:5pt;"> ����: <%=(aCount.getItem(10,0)?aCount.getItem(10,0):0)%> ��.
				  &nbsp;&nbsp; ���: <%=(aCount.getItem(14,0)?aCount.getItem(14,0):0)%> ��. </td>
				  <td class="InnerHead">&nbsp;</td>
			  </tr>
				<tr>
				  <td class="InnerHead">����Ԥ��</td>
				  <td class="InnerMain" style="padding-left:5pt;"> ����: <%=Math.round(aCount.getItem(8,0)/(iNowHour+1)*24)%> ��.�����: <%=Math.round(aCount.getItem(12,0)/(iNowHour+1)*24)%> ��.</td>
				  <td class="InnerHead">�������</td>
				  <td class="InnerMain" style="padding-left:5pt;"><%=aCount.getItem(25,0)%></td>
			  </tr>
      </table></td>
    </tr>
    <tr>
      <td class="OuterFoot"></td>
    </tr>
</table>
	<br>
	<table width="760" border="0" cellpadding="2" cellspacing="1" class="OuterTable">
    <tr>
      <td class="OuterHead"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
          <tr>
            <td width="150" nowrap="nowrap" style="text-align:left">
				<span id="ItemTitle"><font face="webdings">8</font> ���շ�����ϸ</span></td>
            <td nowrap="nowrap" style="text-align:center">
							<marquee id="mq_cc_count_6" width="300" direction=left onmousemove="stop()" onmouseout="start()">
								
							</marquee>
			</td>
            <td width="180" align="right" class="CodingInfo">Coding by Sunrise_Chen. </td>
          </tr>
      </table></td>
    </tr>
    <tr>
      <td><table width="100%"  border="0" cellpadding="0" cellspacing="1" class="InnerTable">
          <tr>
            <td class="InnerHead" width="20%">����ʱ��</td>
            <td width="16%" class="InnerHead">��������Ϣ</td>
            <td class="InnerHead">����ҳ</td>
            <td class="InnerHead">����ҳ</td>
          </tr>
			<%
				var sRefUrl, sPageUrl, sVisitInfo, sLinkType, aClient, sTimeZone;
				if(aVisit!=null){
					for(var i=0;i<=aVisit.ubound(2);++i){
						sRefUrl = aVisit.getItem(2,i);
						sPageUrl = aVisit.getItem(1,i);
						sLinkType = aVisit.getItem(5,i);
						aClient = aVisit.getItem(6,i).split('\n');
						sTimeZone = oTimeZone[aClient[5]] + ' ';
						sVisitInfo = " ------------ " + aVisit.getItem(3,i) + ' ------------ '
							+ '\r\n ����: ' + aVisit.getItem(4,i) 
							+ '\r\n ������ʽ: ' + (sLinkType?sLinkType:'[δ֪]')
							+ '\r\n ����ϵͳ: ' + (aClient[0]?aClient[0]:'[δ֪]')
							+ '\r\n �����: ' + (aClient[1]?aClient[1]:'[δ֪]')
							+ '\r\n �ֱ���: ' + (aClient[2]?aClient[2]:'[δ֪]')
							+ '\r\n ��ɫ���: ' + (aClient[3]?aClient[3]:'[δ֪]') + ' (λ)ɫ'
							+ '\r\n ����: ' + oLanguage[aClient[4]]
							+ '\r\n ʱ��: ' + sTimeZone.substr(0,sTimeZone.indexOf(' '))
							+ '\r\n ------------------------------------------- '
						
			%>
          <tr height="22">
            <td class="InnerMain" align="center" nowrap>
								<%=aVisit.getItem(0,i)%>
			</td>
            <td class="InnerMain" align="left" nowrap title="<%=sVisitInfo%>">
								&nbsp;<%=aVisit.getItem(3,i)%>
			</td>
            <td class="InnerMain" style="padding-left:3pt;padding-right:3pt;" nowrap>
							<div style="width:250;height:14px">
								<a href="<%=sPageUrl%>" title="<%=sPageUrl%>" target="_blank"><%=sPageUrl%></a>
							</div>
			</td>
            <td class="InnerMain" style="padding-left:3pt;padding-right:3pt;" nowrap>
							<div style="width:250;height:14px">
							<%
								if(sRefUrl)
									Response.Write('<a href="'+sRefUrl+'" title="'+sRefUrl+'" target="_blank">'+sRefUrl+'</a>');
								else
									Response.Write('ֱ�������ַ����');
							%>
							</div>
			</td>
          </tr>
					<%
							}
						}
					%>
		  <tr align="right">
		  	<td colspan="4" class="InnerMain">
				<form action="" method="get" style="margin:0" name="formPager">
					�� <%=iRecCount%> ����¼���� <%=iPageCount%> ҳ
					<input type="submit" onClick="form.p.value=1" value="9" <%=iPage==1?"disabled":""%> style="height:18px;font-family:webdings;border:none;" class="InnerMain">
					<input type="submit" onClick="form.p.value=<%=iPage-1%>" value="3" <%=iPage==1?"disabled":""%> style="height:18px;font-family:webdings;border:none;" class="InnerMain">
					<input type="text" name="p" value="<%=iPage%>" style="text-align:center;height:18px;width:24px;border:0px;border-bottom:1px solid black;">
					<input type="submit" onClick="form.p.value=<%=iPage+1%>" value="4" <%=iPage==iPageCount?"disabled":""%> style="height:18px;font-family:webdings;border:none;" class="InnerMain">
					<input type="submit" onClick="form.p.value=<%=iPageCount%>" value=":" <%=iPage==iPageCount?"disabled":""%> style="height:18px;font-family:webdings;border:none;" class="InnerMain">
					&nbsp;&nbsp;
			    </form>			</td>
		  </tr>
      </table></td>
    </tr>
    <tr>
      <td class="OuterFoot"></td>
    </tr>
  </table>
	<% if(Request.QueryString("p").Count>0){ %>
	<script language="javascript">
	//try{
		document.getElementById("formPager").p.select();
	//}catch(e){}
	</script>
	<% } %>
	<!--#include file="../static/bott.htm"-->
</body>
</html>
