<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<%
	var bIsLogined = chkLogin(false);
	iSiteId = 0;	//ע��ҳ��iSiteIdΪ0;
	if(isPostBack()){
		oConn.Open();
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
		if(sNewPwd!=Request.Form("PWD2")+"")
			doAlert("������������벻ͬ������������","javascript:history.back()");

		sSqlString = "select PWD,SiteTitle,SiteIntro,SiteUrl,SiteUrl1,SiteUrl2,SiteMaster,SiteEmail,SiteOpen,UID "
			+ "from t_Site where UID='" + Str4Sql(sUID) + "'";
		var oRs = new ActiveXObject("ADODB.Recordset");
		oRs.open(sSqlString,oConn,1,3);
		if(oRs.EOF) oRs.AddNew();
		else doAlert("�˵�¼���Ѵ��ڣ�����ѡһ����","javascript:history.back()")
		with(oRs){
			fields.item(0).value = sNewPwd;
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
			oConn.Close();
			doAlert("ע��ɹ�","../Supervise/Help.asp");
		}
	}
	
	
	
	
%>
<!--
	COCOON Counter 6
	Develop by Cocoon Studio. [ www.ccopus.com ]
	Product by Sunrise_Chen. (MSN:sunrise_chen@msn.com)
	
	Please don't remove this information, thanks.
-->
<html>
<head>
<title>ע�� - <%=WebTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="../style/main.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.box{border:1px solid #C0C0C0;width:52px;height:19px;clip:rect(0px,181px,18px,0px);overflow:hidden;}
.box2{border:1px solid #F4F4F4;width:50px;height:17px;clip:rect(0px,179px,16px,0px);overflow:hidden;}
select{position:relative;left:-2px;top:-2px;font-size:12px;width:53px;line-height:14px;border:0px;color:#909993;}
-->
</style>
<script language="javascript">
function doSubmit(o){
	sCantBlank = ",UID,PWD1,PWD2,SiteTitle,SiteMaster,SiteEmail,"
	with(o){
		for(var i=0;i<elements.length;++i){
			if(sCantBlank.indexOf(elements[i].name)>0&&elements[i].value==""){
				alert("�뽫����������д������");
				elements[i].focus();
				return false;
			}
		}
	}
	return true;
}
</script>
</head>

<body>
<!--#include file="../supervise/inc_head.asp"-->
<% if(!AllowRegister){ %>
<table width="760" border="0" cellpadding="10" cellspacing="1" class="OuterTable">
<tr>
<td>
ע�Ṧ���ѱ���ֹ������ϵ����Ա��<br>
Email: <a href="mailto:<%=WebMasterEmail%>"><%=WebMasterEmail%></a>
</td>
</tr>
</table>
<% }else{ %>
<form name="form1" method="post" onSubmit="return doSubmit(this)">
<table width="760" border="0" cellpadding="2" cellspacing="1" class="OuterTable">
    <tr>
      <td class="OuterHead"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
          <tr>
            <td nowrap="nowrap" style="text-align:left">
			<span id="ItemTitle"><font face="webdings">8</font> ���û�ע��</span></td>
            <td nowrap="nowrap" style="text-align:left">&nbsp;</td>
            <td width="180" align="right" class="CodingInfo">Coding by Sunrise_Chen. </td>
          </tr>
      </table></td>
    </tr>
    <tr>
      <td><table width="100%"  border="0" cellpadding="0" cellspacing="1" class="InnerTable">
				<tr>
					<td colspan="2" rowspan="12" align="left" class="InnerMain" style="padding:10px;">
						��ӭʹ��COCOON Counter 6<br>
						<br>						<br>
					  ͳ�ƴ��룺
						<iframe src="../Help/GetCodePRO.asp" width="100%" height="120" frameborder="0" style="border: 1px solid buttonface">��ʹ��IE�����</iframe>						
						<BR><BR><BR>
						ͳ����ʽ��<BR>
						<center>
						<script language="javascript">
						var __cc_uid = "c"+"o"+"c"+"o"+"o"+"n";
						var __cc_style = 20;
						</script>
						<script language="javascript" src="http://www.ccopus.com/cccounter6/count.js"></script>
				  </center>				  </td>
					<td class="InnerHead" width="15%">ͳ���ʺ�</td>
					<td class="InnerMain" width="35%" style="padding-left:5pt;">
						<input type="text" name="UID" value="" class="Input">
						<font color=red>*</font>
					</td>
				</tr>
				<tr>
				  <td class="InnerHead">��¼����</td>
				  <td class="InnerMain" style="padding-left:5pt;">
					<input type="password" name="PWD1" value="" class="Input" title="��д������">
					<font color=red>*</font>
				</td>
			  </tr>
			  <tr>
				  <td class="InnerHead">�ظ�����</td>
				  <td class="InnerMain" style="padding-left:5pt;">
					<input type="password" name="PWD2" value="" class="Input" title="�ظ�������">
					<font color=red>*</font>
				  </td>
			  </tr>
				<tr>
				  <td class="InnerHead">վ������</td>
				  <td class="InnerMain" style="padding-left:5pt;">
					<input type="text" name="SiteTitle" value="" class="Input">
					<font color=red>*</font>
          </td>
				</tr>

				<tr>
				  <td class="InnerHead">վ��Url1</td>
				  <td class="InnerMain" style="padding-left:5pt;">
						<input type="text" name="SiteUrl" value="" class="Input">
					</td>
			  </tr>
				<tr>
				  <td class="InnerHead">վ��Url2</td>
				  <td class="InnerMain" style="padding-left:5pt;">
						<input type="text" name="SiteUrl1" value="" class="Input">
					</td>
			  </tr>
				<tr>
				  <td class="InnerHead">վ��Url3</td>
				  <td class="InnerMain" style="padding-left:5pt;">
						<input type="text" name="SiteUrl2" value="" class="Input">
					</td>
			  </tr>
				<tr>
				  <td class="InnerHead">վ���ǳ�</td>
				  <td class="InnerMain" style="padding-left:5pt;">
						<input type="text" name="SiteMaster" value="" class="Input">
						<font color=red>*</font>
					</td>
			  </tr>
				<tr>
				  <td class="InnerHead">վ������</td>
				  <td class="InnerMain" style="padding-left:5pt;">
					<input type="text" name="SiteEmail" value="" class="Input">
					<font color=red>*</font>
				</td>
			  </tr>
				<tr>
				  <td class="InnerHead">վ����</td>
				  <td rowspan="2" valign="top" class="InnerMain" style="padding-left:5pt;padding-top:3pt">
						<textarea name="SiteIntro" class="Input" wrap="VIRTUAL"></textarea>
					</td>
			  </tr>
				<tr>
				  <td class="InnerHead">&nbsp;</td>
			  </tr>
				<tr>
				  <td class="InnerHead">�������</td>
				  <td class="InnerMain" style="padding-left:5pt;">
						<input type="radio" name="SiteOpen" value="1" id="radioSiteOpen_0" checked><label for="radioSiteOpen_0">����</label>
						<input type="radio" name="SiteOpen" value="0" id="radioSiteOpen_1" ><label for="radioSiteOpen_1">������</label>
					</td>
			  </tr>
				<tr>
				  <td colspan="2" class="InnerHead">&nbsp;</td>
				  <td colspan="2" class="InnerMain" style="text-align:right">
						<input type="submit" value="����" class="GenButton">
					&nbsp;
					<input type="Reset" value="ȡ��" class="GenButton">
					&nbsp;</td>
			  </tr>
      </table></td>
    </tr>
    <tr>
      <td class="OuterFoot"></td>
    </tr>
  </table>
</form>
<% }	//�Ƿ��ֹע�� %>
<!--#include file="../static/bott.htm"-->
</body>
</html>
