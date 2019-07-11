<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<%
	oConn.Open();
	
	var bIsLogined = chkLogin(true);
	
	if(isPostBack()){
		var sUID = Request.Form("UID")+"";
		var sPwd = Request.Form("PWD")+"";
		var sNewPwd = Request.Form("PWD1")+"";
		var sSiteTitle = Request.Form("SiteTitle")+"";
		var sSiteIntro = Request.Form("SiteIntro")+"";
		var sSiteUrl = Request.Form("SiteUrl")+"";
		var sSiteUrl1 = Request.Form("SiteUrl1")+"";
		var sSiteUrl2 = Request.Form("SiteUrl2")+"";
		var sSiteMaster = Request.Form("SiteMaster")+"";
		var sSiteEmail = Request.Form("SiteEmail")+"";
		var sSiteOpen = Request.Form("SiteOpen")+"";
		sSqlString = "select PWD,SiteTitle,SiteIntro,SiteUrl,SiteUrl1,SiteUrl2,SiteMaster,SiteEmail,SiteOpen,UID "
			+ "from t_Site where SiteId=" + iSiteId;
		var oRs = new ActiveXObject("ADODB.Recordset");
		oRs.open(sSqlString,oConn,1,3);
		if(oRs.EOF){ Response.End() }	//Error
		if(oRs.fields.item(0).value!=sPwd){
			doAlert("密码错误","javascript:history.back()");
		}
		with(oRs){
			if(sNewPwd){
				if(sNewPwd==Request.Form("PWD2")+"")
					fields.item(0).value = sNewPwd;
				else
					doAlert("两次输入的密码不同，请重新输入","javascript:history.back()");
			}
			fields.item(1).value = sSiteTitle;
			fields.item(2).value = sSiteIntro;
			fields.item(3).value = sSiteUrl;
			fields.item(4).value = sSiteUrl1;
			fields.item(5).value = sSiteUrl2;
			fields.item(6).value = sSiteMaster;
			fields.item(7).value = sSiteEmail;
			fields.item(8).value = sSiteOpen;
			fields.item(9).value = sUID;
			update();
			doAlert("修改成功");
		}
	}
	
	sSqlString = "select top 1 UID,SiteTitle,SiteIntro,SiteUrl,SiteUrl1,SiteUrl2,SiteMaster,SiteEmail,SiteOpen "
		+ "from t_Site where SiteId=" + iSiteId;
	var oRs = oConn.Execute( sSqlString );
	if(oRs.EOF){
		oConn.Close(); Response.End();
	}
	var aInfo = oRs.GetRows();
	
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
<title>资料修改 - <%=WebTitle%></title>
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
<form name="form1" method="post">
<table width="760" border="0" cellpadding="2" cellspacing="1" class="OuterTable">
    <tr>
      <td class="OuterHead"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
          <tr>
            <td nowrap="nowrap" style="text-align:left">
			<span id="ItemTitle"><font face="webdings">8</font> 修改资料 -- <%=aInfo.getItem(3,0).replace("http://","")%></span></td>
            <td nowrap="nowrap" style="text-align:left">&nbsp;</td>
            <td width="180" align="right" class="CodingInfo">Coding by Sunrise_Chen. </td>
          </tr>
      </table></td>
    </tr>
    <tr>
      <td><table width="100%"  border="0" cellpadding="0" cellspacing="1" class="InnerTable">
				<tr>
					<td colspan="2" rowspan="12" align="center" class="InnerMain"><p>Welcome to our COCOON sTudio.</p>
				    <p>We will show you the place where dreams and life become one ...  </p></td>
					<td class="InnerHead" width="15%">站点名称</td>
					<td class="InnerMain" width="35%" style="padding-left:5pt;">
						<input type="text" name="SiteTitle" value="<%=aInfo.getItem(1,0)%>" class="Input">
					</td>
				</tr>
				<tr>
				  <td class="InnerHead">站点Url1</td>
				  <td class="InnerMain" style="padding-left:5pt;">
						<input type="text" name="SiteUrl" value="<%=aInfo.getItem(3,0)%>" class="Input">
					</td>
			  </tr>
				<tr>
				  <td class="InnerHead">站点Url2</td>
				  <td class="InnerMain" style="padding-left:5pt;">
						<input type="text" name="SiteUrl1" value="<%=aInfo.getItem(4,0)%>" class="Input">
					</td>
			  </tr>
				<tr>
				  <td class="InnerHead">站点Url3</td>
				  <td class="InnerMain" style="padding-left:5pt;">
						<input type="text" name="SiteUrl2" value="<%=aInfo.getItem(5,0)%>" class="Input">
					</td>
			  </tr>
				<tr>
				  <td class="InnerHead">站长昵称</td>
				  <td class="InnerMain" style="padding-left:5pt;">
						<input type="text" name="SiteMaster" value="<%=aInfo.getItem(6,0)%>" class="Input">
					</td>
			  </tr>
				<tr>
				  <td class="InnerHead">站长信箱</td>
				  <td class="InnerMain" style="padding-left:5pt;">
						<input type="text" name="SiteEmail" value="<%=aInfo.getItem(7,0)%>" class="Input">
					</td>
			  </tr>
				<tr>
				  <td class="InnerHead">统计帐号</td>
				  <td class="InnerMain" style="padding-left:5pt;">
						<input type="text" name="UID" value="<%=aInfo.getItem(0,0)%>" class="Input">
					</td>
			  </tr>
				<tr>
				  <td class="InnerHead">管理密码</td>
				  <td class="InnerMain" style="padding-left:5pt;">
						<input type="password" name="PWD" value="" class="Input" title="填写原密码">
					</td>
			  </tr>
				<tr>
				  <td class="InnerHead">修改密码</td>
				  <td class="InnerMain" style="padding-left:5pt;">
						<input type="password" name="PWD1" value="" class="Input" style="width:119px;" title="填写新密码，如果不需要修改密码，请置空">
						<input type="password" name="PWD2" value="" class="Input" style="width:118px;" title="重复新密码">
					</td>
			  </tr>
				<tr>
				  <td class="InnerHead">站点简介</td>
				  <td rowspan="2" valign="top" class="InnerMain" style="padding-left:5pt;padding-top:3pt">
						<textarea name="SiteIntro" class="Input"><%=aInfo.getItem(2,0)%></textarea>
					</td>
			  </tr>
				<tr>
				  <td class="InnerHead">&nbsp;</td>
			  </tr>
				<tr>
				  <td class="InnerHead">公开结果</td>
				  <td class="InnerMain" style="padding-left:5pt;">
						<input type="radio" name="SiteOpen" value="1" id="radioSiteOpen_0" <%=aInfo.getItem(8,0)?"checked":""%>><label for="radioSiteOpen_0">公开</label>
						<input type="radio" name="SiteOpen" value="0" id="radioSiteOpen_1" <%=aInfo.getItem(8,0)?"":"checked"%>><label for="radioSiteOpen_1">不公开</label>
					</td>
			  </tr>
				<tr>
				  <td colspan="2" class="InnerHead">&nbsp;</td>
				  <td colspan="2" class="InnerMain" style="text-align:right">
						<input type="submit" value="保存" class="GenButton">
					&nbsp;
					<input type="Reset" value="取消" class="GenButton">
					&nbsp;</td>
			  </tr>
      </table></td>
    </tr>
    <tr>
      <td class="OuterFoot"></td>
    </tr>
  </table>
</form>
<!--#include file="../static/bott.htm"-->
</body>
</html>
