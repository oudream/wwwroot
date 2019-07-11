<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<SCRIPT LANGUAGE="JavaScript">
<!--
	  function rable(value)
	  { 
	        if  (value=="ddns"){
			document.domainform.d4.value="";
			document.domainform.d5.value="";
			document.domainform.d4.disabled=true;
			document.domainform.d5.disabled=true;
			document.domainform.d1.disabled=false;
			document.domainform.d2.disabled=false;
	  }
	   
	     if  (value=="sdns"){
			document.domainform.d4.value="ns1.dnyes.com";
			document.domainform.d5.value="ns2.dnyes.com";
			document.domainform.d4.disabled=false;
			document.domainform.d5.disabled=false;
			document.domainform.d1.value="";
			document.domainform.d2.value="";
			document.domainform.d1.disabled=true;
			document.domainform.d2.disabled=true;
						
	  }
	}

function fill()
{
  domainform.danc.value = domainform.doceoc.value;
  domainform.dane.value = domainform.doceoe.value;
  domainform.daoc.value = domainform.donc.value;
  domainform.daoe.value = domainform.done.value;
  domainform.das.value = domainform.dos.value;
  domainform.dacc.value = domainform.docc.value;
  domainform.dace.value = domainform.doce.value;
  domainform.daac.value = domainform.doac.value;
  domainform.daae.value = domainform.doae.value;
  domainform.daz.value = domainform.doz.value;
}

function doValidate()
{
	if(! isChinese(domainform.donc.value)) {
		alert("单位中文名称中不能没有中文！");
		domainform.donc.focus();
		return false;
	}
	if(! isEnglish(domainform.done.value)) {
		alert("单位英文名称中不可以含有中文！");
		domainform.done.focus();
		return false;
	}
	if(! isChinese(domainform.doceoc.value)) {
		alert("单位负责人中文名不合法！");
		domainform.doceoc.focus();
		return false;
	}
	if(! isName(domainform.doceoe.value)) {
		alert("单位负责人英文名不合法！");
		domainform.doceoe.focus();
		return false;
	}
	if(domainform.dos.value.indexOf("NULL") != -1) {
		alert("请选择一个省份！");
		domainform.dos.focus();
		return false;
	}
	if(! isChinese(domainform.docc.value)) {
		alert("城市中文名称中不能没有中文！");
		domainform.docc.focus();
		return false;
	}
	if(! isEnglish(domainform.doce.value)) {
		alert("城市英文名称中不可以含有中文！");
		domainform.doce.focus();
		return false;
	}
	if(! isChinese(domainform.doac.value)) {
		alert("街道中文名称中不能没有中文！");
		domainform.doac.focus();
		return false;
	}
	if(! isEnglish(domainform.doae.value)) {
		alert("街道英文名称中不可以含有中文！");
		domainform.doae.focus();
		return false;
	}
	if(! isNumber(domainform.doz.value)) {
		alert("邮政编码不合法！");
		domainform.doz.focus();
		return false;
	}
	if(! isChinese(domainform.danc.value)) {
		alert("管理联系人中文名不合法！");
		domainform.danc.focus();
		return false;
	}
	if(! isName(domainform.dane.value)) {
		alert("管理联系人英文名不合法！");
		domainform.dane.focus();
		return false;
	}
	if(! isChinese(domainform.daoc.value)) {
		alert("单位中文名称中不能没有中文！");
		domainform.daoc.focus();
		return false;
	}
	if(! isEnglish(domainform.daoe.value)) {
		alert("单位英文名称中不可以含有中文！");
		domainform.daoe.focus();
		return false;
	}
	if(domainform.das.value.indexOf("NULL") != -1) {
		alert("请选择一个省份！");
		domainform.das.focus();
		return false;
	}
	if(! isChinese(domainform.dacc.value)) {
		alert("城市中文名称中不能没有中文！");
		domainform.dacc.focus();
		return false;
	}
	if(! isEnglish(domainform.dace.value)) {
		alert("城市英文名称中不可以含有中文！");
		domainform.dace.focus();
		return false;
	}
	if(! isChinese(domainform.daac.value)) {
		alert("街道中文名称中不能没有中文！");
		domainform.daac.focus();
		return false;
	}
	if(! isEnglish(domainform.daae.value)) {
		alert("街道英文名称中不可以含有中文！");
		domainform.daae.focus();
		return false;
	}
	if(! isNumber(domainform.daz.value)) {
		alert("邮政编码不合法！");
		domainform.daz.focus();
		return false;
	}
	if(! isNumber(domainform.dapi.value)) {
		alert("管理联系人电话国家区号不合法！");
		domainform.dapi.focus();
		return false;
	}
	if(! isNumber(domainform.dapa.value)) {
		alert("管理联系人电话地区区号不合法！");
		domainform.dapa.focus();
		return false;
	}
	if(! isNumber(domainform.dapn.value)) {
		alert("管理联系人电话号码不合法！");
		domainform.dapn.focus();
		return false;
	}
	if(domainform.dape.value.length > 0) {
		if(! isNumber(domainform.dape.value)) {
			alert("管理联系人电话分机号码不合法！");
			domainform.dape.focus();
			return false;
		}
	}
	if(! isNumber(domainform.dafi.value)) {
		alert("管理联系人传真国家区号不合法！");
		domainform.dafi.focus();
		return false;
	}
	if(! isNumber(domainform.dafa.value)) {
		alert("管理联系人传真地区区号不合法！");
		domainform.dafa.focus();
		return false;
	}
	if(! isNumber(domainform.dafn.value)) {
		alert("管理联系人传真号码不合法！");
		domainform.dafn.focus();
		return false;
	}
	if(! isMail(domainform.daem.value)) {
		alert("管理联系人电子邮件不合法！");
		domainform.daem.focus();
		return false;
	}
	return true;
}

function isChinese(name)
{
	if(name.length == 0)
		return false;
	for(i = 0; i < name.length; i++) {
		if(name.charCodeAt(i) > 128)
			return true;
	}
	return false;
}

function isEnglish(name)
{
	if(name.length == 0)
		return false;
	for(i = 0; i < name.length; i++) {
		if(name.charCodeAt(i) > 128)
			return false;
	}
	return true;
}

function isName(name)
{
	if(! isEnglish(name))
		return false;
	i = name.indexOf(" ");
	if(i == -1)
		return false;
	if(i == name.length)
		return false;
	return true;
}

function isNumber(name)
{
	if(name.length == 0)
		return false;
	for(i = 0; i < name.length; i++) {
		if(name.charAt(i) < "0" || name.charAt(i) > "9")
			return false;
	}
	return true;
}

function isMail(name)
{
	if(! isEnglish(name))
		return false;
	i = name.indexOf("@");
	j = name.lastIndexOf("@");
	if(i == -1)
		return false;
	if(i != j)
		return false;
	if(i == name.length)
		return false;
	return true;
}
//-->
</SCRIPT>
<%
domain=request("domain")
after=request("after")
domaintypestr=request("domaintypestr")
session("domainurl")="domain.asp" & "?alldomain=" & after & "," & domain & "," & domaintypestr
if session("y_it_uid")="" then
response.redirect "error.asp?err=012"
response.end
end if
%>
<%
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from user where uid='"&session("y_it_uid")&"'"
rs.open sql,conn,1,1
%>
<html>
<head>
<title>信网-数信科技 域名注册 A记录、MX记录、URL转发可自行修改，自助MyDNS功能，免费赠送20个二级域名、自由免费修改域名信息，让您成为域名的真正拥有者</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="www.dnyes.com-domain,域名注册,whois查询,A记录（IP指向）,MX记录（邮件记录）,URL转发（网址转发）,强大的自助MyDNS功能，免费赠送20个二级域名、注册人可以自由免费修改域名信息!">
<meta name="description" content="www.dnyes.com-domain,域名注册,whois查询,A记录（IP指向）,MX记录（邮件记录）,URL转发（网址转发）,强大的自助MyDNS功能，免费赠送20个二级域名、注册人可以自由免费修改域名信息!">
<link rel="stylesheet" href="css.css" type="text/css">
</head>
<body leftmargin="0" topmargin="0">
<!--#include file="top.asp" -->
<table width="776" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="10" background="images/dnyes-bg-left-and-right.gif"><img src="images/dnyes-bg-left-and-right.gif" width="10" height="1"></td>
    <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td colspan="5"><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="1" background="images/2x2.gif"><img src="images/1x2-black.gif" width="1" height="2"></td>
                <td><img src="images/1-x.gif" width="754" height="56"></td>
                <td width="1" background="images/1x2-black.gif"><img src="images/1x2-black.gif" width="1" height="2"></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td valign="top"><TABLE width=100% border=0 cellPadding=0 cellSpacing=0>
              <TBODY>
                <TR> 
                  <TD align="center" 
                vAlign=top 
                style="BORDER-left: #000000 1px solid; BORDER-RIGHT: #000000 1px solid"> 
				
===<b>域名注册申请表</b>===
<table border="0" cellspacing="0" cellpadding="1" align="center" width="98%">
    <tr>
      <td>
        <table border="0" cellpadding="3" cellspacing="1" bgcolor="#000000">
<form class="form" method="POST" action="domainregsucc.asp" name="domainform" onsubmit="return doValidate()">
                              <tr bgcolor="#FFFFFF"> 
                                <td colspan="3"> 填写域名注册申请表 
                                </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">1. 域名</td>
                                <td colspan="2"><font color=red><%=domain%><%=after%></td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%"><div align="left"></div></td>
                                <td colspan="2">&nbsp; </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">2. 注册单位</td>
                                <td colspan="2" >　</td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%"> 2.1 单位名称：（中文）<br>
                                  如个人注册，填写姓名即可 </td>
                                <td colspan="2">  
                                  <input name="donc" size="40" maxlength="60" value="<%=rs("namez")%>">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%"> 2.2 单位名称：（英文）<br>
                                  如个人注册，填写姓名即可 </td>
                                <td colspan="2">  
                                  <input name="done" size="40" maxlength="80" value="<%=rs("namee")%>">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">2.3 单位负责人：（中文）</td>
                                <td colspan="2"> 
                                  <input name="doceoc" type="text" value="<%=rs("contactz")%>" size="20" maxlength="20">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">2.4 单位负责人：（英文）<br>
                                  英文的姓和名之间必须有空格间隔</td>
                                <td colspan="2"> 
                                  <input name="doceoe" type="text" value="<%=rs("contacte")%>" size="20" maxlength="30">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">2.5 省份：</td>
                                <td colspan="2"> <select name="dos" size="1"> 
                                    <option value=<%=rs("shengz")%> selected><%=rs("shengz")%></option>
                                    <option value="Beijing">北京</option>
                                    <option value="Hongkong">香港</option>
                                    <option value="Aomen">澳门</option>
                                    <option value="Taiwan">台湾</option>
                                    <option value="Shanghai">上海</option>
                                    <option value="Guangdong">广东</option>
                                    <option value="Shandong">山东</option>
                                    <option value="Sichuan">四川</option>
                                    <option value="Fujian">福建</option>
                                    <option value="Jiangsu">江苏</option>
                                    <option value="Zhejiang">浙江</option>
                                    <option value="Tianjin">天津</option>
                                    <option value="Chongqing">重庆</option>
                                    <option value="Hebei">河北</option>
                                    <option value="Henan">河南</option>
                                    <option value="Heilongjiang">黑龙江</option>
                                    <option value="Jinlin">吉林</option>
                                    <option value="Liaoning">辽宁</option>
                                    <option value="Neimenggu">内蒙古</option>
                                    <option value="Hainan">海南</option>
                                    <option value="Shanxi">山西</option>
                                    <option value="Shanxi3">陕西</option>
                                    <option value="Anhui">安徽</option>
                                    <option value="Jiangxi">江西</option>
                                    <option value="Gansu">甘肃</option>
                                    <option value="Xinjiang">新疆</option>
                                    <option value="Hubei">湖北</option>
                                    <option value="Hunan">湖南</option>
                                    <option value="Yunnan">云南</option>
                                    <option value="Guangxi">广西</option>
                                    <option value="Ningxia">宁夏</option>
                                    <option value="Guizhou">贵州</option>
                                    <option value="Qinghai">青海</option>
                                    <option value="Xizang">西藏</option>
                                  </select> </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">2.6 城市：（中文）</td>
                                <td colspan="2">  
                                  <input maxlength="20" name="docc" size="20" value="<%=rs("cityz")%>">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">2.7 城市：（英文）</td>
                                <td colspan="2">  
                                  <input maxlength="30" name="doce" size="20" value="<%=rs("citye")%>">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">2.8 街道地址：（中文）</td>
                                <td colspan="2">  
                                  <input maxlength="100" name="doac" size="40" value="<%=rs("dizhiz")%>">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">2.9 街道地址：（英文）</td>
                                <td colspan="2">  
                                  <input maxlength="120" name="doae" size="40" value="<%=rs("dizhie")%>">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">2.10 邮政编码：</td>
                                <td colspan="2">  
                                  <input maxlength="8" name="doz" size="10" value="<%=rs("postnumber")%>">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%"><div align="left"></div></td>
                                <td colspan="2">&nbsp; </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%" rowspan="2">3.管理联系人</td>
                                <td width="35%"> 
                                  <input type="radio" value="V1" name="R1" onClick="fill()">
                                  使用注册单位的信息 </td>
                                <td width="34%"> 
                                  <input type="radio" value="V2" name="R1" checked>
                                  填写具体信息</td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td colspan="2"> <font color="#FF0000">如果要注册国内域名，请保持管理联系人的信息与注册单位信息一致，否则注册国内域名将会失败。 
                                </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">3.1 姓名：（中文）</td>
                                <td colspan="2">  
                                  <input name="danc" size="20" maxlength="20">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%"> 3.2 姓名：（英文）<br>
                                  英文的姓和名之间必须有空格间隔 </td>
                                <td colspan="2">  
                                  <input name="dane" size="20" maxlength="30">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">3.3 单位名称：（中文）</td>
                                <td colspan="2">  
                                  <input maxlength="80" name="daoc" size="40">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">3.4 单位名称：（英文）</td>
                                <td colspan="2">  
                                  <input maxlength="100" name="daoe" size="40">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">3.5 省份：</td>
                                <td colspan="2">  
                                  <select name="das" size="1">
                                    <option value=<%=rs("shengz")%> selected><%=rs("shengz")%></option>
                                    <option value="Beijing">北京</option>
                                    <option value="Hongkong">香港</option>
                                    <option value="Aomen">澳门</option>
                                    <option value="Taiwan">台湾</option>
                                    <option value="Shanghai">上海</option>
                                    <option value="Guangdong">广东</option>
                                    <option value="Shandong">山东</option>
                                    <option value="Sichuan">四川</option>
                                    <option value="Fujian">福建</option>
                                    <option value="Jiangsu">江苏</option>
                                    <option value="Zhejiang">浙江</option>
                                    <option value="Tianjin">天津</option>
                                    <option value="Chongqing">重庆</option>
                                    <option value="Hebei">河北</option>
                                    <option value="Henan">河南</option>
                                    <option value="Heilongjiang">黑龙江</option>
                                    <option value="Jinlin">吉林</option>
                                    <option value="Liaoning">辽宁</option>
                                    <option value="Neimenggu">内蒙古</option>
                                    <option value="Hainan">海南</option>
                                    <option value="Shanxi">山西</option>
                                    <option value="Shanxi3">陕西</option>
                                    <option value="Anhui">安徽</option>
                                    <option value="Jiangxi">江西</option>
                                    <option value="Gansu">甘肃</option>
                                    <option value="Xinjiang">新疆</option>
                                    <option value="Hubei">湖北</option>
                                    <option value="Hunan">湖南</option>
                                    <option value="Yunnan">云南</option>
                                    <option value="Guangxi">广西</option>
                                    <option value="Ningxia">宁夏</option>
                                    <option value="Guizhou">贵州</option>
                                    <option value="Qinghai">青海</option>
                                    <option value="Xizang">西藏</option>
                                  </select>
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">3.6 城市：（中文）</td>
                                <td colspan="2">  
                                  <input maxlength="20" name="dacc" size="20">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">3.7 城市：（英文）</td>
                                <td colspan="2">  
                                  <input maxlength="30" name="dace" size="20">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">3.8 街道地址：（中文）</td>
                                <td colspan="2">  
                                  <input maxlength="100" name="daac" size="40">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">3.9 街道地址：（英文）</td>
                                <td colspan="2">  
                                  <input maxlength="120" name="daae" size="40">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">3.10 邮政编码：</td>
                                <td colspan="2">  
                                  <input maxlength="8" name="daz" size="10">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">3.11 电话：</td>
                                <td colspan="2"> 
<%
if rs("tel")<>"" then tel=split(rs("tel"),"-")
%>								 
                                  <input name="dapi" size="5" maxlength="7" value="<%=tel(0)%>">
                                  <b>-</b> 
                                  <input name="dapa" size="7" maxlength="7" value="<%=tel(1)%>">
                                  <b>-</b> 
                                  <input name="dapn" size="15" maxlength="15" value="<%=tel(2)%>">
                                  <b>-</b> 
                                  <input maxlength="7" name="dape" size="7" value="">
                                  <br>
                                  国家区号 - 地区区号 -&nbsp;&nbsp;&nbsp;&nbsp; 电话号码&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- 
                                  分机号码 </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">3.12 传真（如没有可填电话号码）：</td>
                                <td colspan="2">  
                                  <input name="dafi" size="5" maxlength="7" value="86">
                                  <b>-</b> 
                                  <input name="dafa" size="7" maxlength="7">
                                  <b>-</b> 
                                  <input name="dafn" size="15" maxlength="15">
                                  <input type="hidden" maxlength="7" name="dafe" size="7">
                                  <br>
                                  国家区号 - 地区区号 -&nbsp;&nbsp;&nbsp;&nbsp; 传真号码</td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">3.13 电子邮件：</td>
                                <td colspan="2">  
                                  <input name="daem" size="40" maxlength="40" value="<%=rs("email")%>">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td colspan="3"><div align="left"></div>
                                  &nbsp; </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%" rowspan="4" align="center"><p> 
                                    域名服务器 ：</p>
                                  <p>（辅域名服务器没有设置可以不填）</p></td>
                                <td colspan="2">&nbsp; &nbsp; 
                                   
                                  <input type="radio" name="radiobutton"  checked  value="dns3"  onClick='rable("sdns")' style="border-style: none">
                                  使用默认DNS服务器</td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td> DNS1 
                                  <input name="d4" size="30" maxlength="40" value="ns1.dnyes.com">
                                  </td>
                                <td> DNS2 
                                  <input name="d5" size="30" maxlength="40" value="ns2.dnyes.com">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td colspan="2"> &nbsp; 
                                  &nbsp; 
                                  <input type="radio" name="radiobutton" value="dns1" onClick='rable("ddns")' style="border-style: none">
                                  您填写的DNS必须是注册过的有效的DNS </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td> DNS1 
                                  <input name="d1" size="30" maxlength="40" value="" disabled>
                                   &nbsp; </td>
                                <td> DNS2 
                                  <input name="d2" size="30" maxlength="40" value="" disabled>
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td colspan="3">&nbsp; 特别提醒您：为了使您更快捷地注册域名，此处只需填写注册人信息，其他管理联系人、技术联系人、付费联系人信息请您在注册成功后登陆会员服务系统免费更改即可。</td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td colspan="3" align="center">  
                                  <input class=botton type="submit" name="B1" value="填写好了，到下一步">
                                  <input class=botton type="reset" value="全部清除，重新填写" name="B2">
                                  <input type="hidden" name="domain" value="<%=domain%>">
                                  <input type="hidden" name="after" value="<%=after%>">
                                  </td>
                              </tr>
</form>
                            </table>
      </td>
    </tr>
  </table>
<%
rs.close
set rs=nothing
%>

<!--FORM结束-->
				</TD>
                </TR>
              </TBODY>
            </TABLE></td>
        </tr>
      </table></td>
    <td width="10" background="images/dnyes-bg-left-and-right.gif"><img src="images/dnyes-bg-left-and-right.gif" width="10" height="1"></td>
  </tr>
</table>
<!--#include file="copyright2.asp"-->
</body>
</html>