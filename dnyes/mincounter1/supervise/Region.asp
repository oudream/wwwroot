<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<%
	oConn.Open();
	
	var bIsLogined = chkLogin(true);
	
	var iWidth = 160;
	
	var aClientRegion = new Array();
	var iRegionSum=0, iRegionMax=0,iTmpValue=0;
	
	var sOrderBy = '';
	var sSearchRegion = Request.Form('SearchRegion')+'';
	var sSearchRegion1 = Request.Form('SearchRegion1')+'';
	if(sSearchRegion) sOrderBy += 'instr(Region,"'+sSearchRegion+'") desc,'
	if(sSearchRegion1) sOrderBy += 'instr(Region,"'+sSearchRegion1+'") desc,'
	
	sSqlString = "select top 20 Region,RegionCount,RegionIp,CreateTime from t_Region where SiteId=" + iSiteId + ' order by '+sOrderBy+'RegionCount desc';
	var oRs = oConn.Execute( sSqlString );
	if(!oRs.EOF){
		var aRs = oRs.GetRows(20);
		for(var i=0;i<=aRs.ubound(2);++i){
			iTmpValue = aRs.getItem(1,i);
			iRegionSum += iTmpValue;
			iRegionMax = (iTmpValue>iRegionMax?iTmpValue:iRegionMax);
			aClientRegion[aClientRegion.length] = Array(aRs.getItem(0,i),iTmpValue,aRs.getItem(2,i),formatDateTime(new Date(aRs.getItem(3,i)),0));
		}
		sSqlString = "select sum(RegionCount) from t_Region where SiteId=" + iSiteId
		iRegionSum = oConn.Execute(sSqlString)(0).Value;
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
<title>����ͳ�� - <%=WebTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="../style/main.css" rel="stylesheet" type="text/css" />
</head>

<body>
<!--#include file="inc_head.asp"-->
	<table width="760" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="45%" align="left" valign="top"><p>
				<fieldset style="padding:3px">
					<legend>��ѯ</legend>
					<form name="formSearchRegion" method="post" class="OuterHead" style="height:24px">
					<table class="SubHeader"><tr><td>
					������
						<input type="text" name="SearchRegion" class="ThinBorderText" style="width:85px">
						��
						<input type="text" name="SearchRegion1" class="ThinBorderText" style="width:85px">
						<input type="button" value="��ѯ" onClick="form.submit();" class="button">
					</td></tr></table>
					</form>
				</fieldset>
      <p></p>
        <p>*˵���� �������ݽ�Ϊ��֪���ݣ�δ֪���ݲ���ͳ��֮�С�</p>
      <p>&nbsp;</p>
      <p>�������ԡ����ޱ�� wzpg.com������л���ǡ�</p>
      <p>COCOON Counter �������ڸ������ݿ��Ա�֤��Ϣ��׼ȷ�ԣ������ӵ��COCOON Counter�Ŀ��������붨�ڵ�������վ������ѯ�Ƿ����µĸ��¡�</p>
      <p align="right">-- Sunrise_Chen </p></td>
      <td width="55%" align="right" valign="top">
        <table width="400" border="0" cellpadding="2" cellspacing="1" class="OuterTable">
          <tr>
            <td class="OuterHead"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
                <tr>
                  <td nowrap="nowrap" style="text-align:left">
					<span id="ItemTitle"><font face="webdings">8</font> ����ͳ�� (ǰ20λ)</span></td>
                  <td align="right" class="CodingInfo">Coding by Sunrise_Chen. </td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td><table width="100%"  border="0" cellpadding="0" cellspacing="0" class="OuterTable">
    <%
			for(var i=0;i<aClientRegion.length;++i){
				iTmpValue = aClientRegion[i][1];
				var iWidthPersent = Math.round(iTmpValue/iRegionSum*(iWidth-5)*100)/100+1;
				var iTotalPersent = Math.round(iTmpValue/iRegionSum*10000)/100;
				Response.Write(
					  '<tr>'
					+ '<td class="InnerHead" height="18" style="padding-left:20px;text-align:left">'+aClientRegion[i][0]+'</td>'
					+ '<td class="InnerMain" height="18" style="width:120px;text-align:center">'+aClientRegion[i][3]+'</td>'
					+ '<td width="'+iWidth+'" class="InnerMain" style="text-align:left" title="'+aClientRegion[i][0]+'\n'+iTmpValue+' ('+iTotalPersent+'%)\n  ------------\n������: '+aClientRegion[i][2]+'\n��������: '+aClientRegion[i][3]+'">'
					+ '<span class="'+(iTmpValue==iRegionMax?'HotColumnH':'ColumnH')+'" style="width:'+iWidthPersent+'px">'+(iTotalPersent>10?iTmpValue+'&nbsp;':'')+'</span>&nbsp;'
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
