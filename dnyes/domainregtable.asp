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
		alert("��λ���������в���û�����ģ�");
		domainform.donc.focus();
		return false;
	}
	if(! isEnglish(domainform.done.value)) {
		alert("��λӢ�������в����Ժ������ģ�");
		domainform.done.focus();
		return false;
	}
	if(! isChinese(domainform.doceoc.value)) {
		alert("��λ���������������Ϸ���");
		domainform.doceoc.focus();
		return false;
	}
	if(! isName(domainform.doceoe.value)) {
		alert("��λ������Ӣ�������Ϸ���");
		domainform.doceoe.focus();
		return false;
	}
	if(domainform.dos.value.indexOf("NULL") != -1) {
		alert("��ѡ��һ��ʡ�ݣ�");
		domainform.dos.focus();
		return false;
	}
	if(! isChinese(domainform.docc.value)) {
		alert("�������������в���û�����ģ�");
		domainform.docc.focus();
		return false;
	}
	if(! isEnglish(domainform.doce.value)) {
		alert("����Ӣ�������в����Ժ������ģ�");
		domainform.doce.focus();
		return false;
	}
	if(! isChinese(domainform.doac.value)) {
		alert("�ֵ����������в���û�����ģ�");
		domainform.doac.focus();
		return false;
	}
	if(! isEnglish(domainform.doae.value)) {
		alert("�ֵ�Ӣ�������в����Ժ������ģ�");
		domainform.doae.focus();
		return false;
	}
	if(! isNumber(domainform.doz.value)) {
		alert("�������벻�Ϸ���");
		domainform.doz.focus();
		return false;
	}
	if(! isChinese(domainform.danc.value)) {
		alert("������ϵ�����������Ϸ���");
		domainform.danc.focus();
		return false;
	}
	if(! isName(domainform.dane.value)) {
		alert("������ϵ��Ӣ�������Ϸ���");
		domainform.dane.focus();
		return false;
	}
	if(! isChinese(domainform.daoc.value)) {
		alert("��λ���������в���û�����ģ�");
		domainform.daoc.focus();
		return false;
	}
	if(! isEnglish(domainform.daoe.value)) {
		alert("��λӢ�������в����Ժ������ģ�");
		domainform.daoe.focus();
		return false;
	}
	if(domainform.das.value.indexOf("NULL") != -1) {
		alert("��ѡ��һ��ʡ�ݣ�");
		domainform.das.focus();
		return false;
	}
	if(! isChinese(domainform.dacc.value)) {
		alert("�������������в���û�����ģ�");
		domainform.dacc.focus();
		return false;
	}
	if(! isEnglish(domainform.dace.value)) {
		alert("����Ӣ�������в����Ժ������ģ�");
		domainform.dace.focus();
		return false;
	}
	if(! isChinese(domainform.daac.value)) {
		alert("�ֵ����������в���û�����ģ�");
		domainform.daac.focus();
		return false;
	}
	if(! isEnglish(domainform.daae.value)) {
		alert("�ֵ�Ӣ�������в����Ժ������ģ�");
		domainform.daae.focus();
		return false;
	}
	if(! isNumber(domainform.daz.value)) {
		alert("�������벻�Ϸ���");
		domainform.daz.focus();
		return false;
	}
	if(! isNumber(domainform.dapi.value)) {
		alert("������ϵ�˵绰�������Ų��Ϸ���");
		domainform.dapi.focus();
		return false;
	}
	if(! isNumber(domainform.dapa.value)) {
		alert("������ϵ�˵绰�������Ų��Ϸ���");
		domainform.dapa.focus();
		return false;
	}
	if(! isNumber(domainform.dapn.value)) {
		alert("������ϵ�˵绰���벻�Ϸ���");
		domainform.dapn.focus();
		return false;
	}
	if(domainform.dape.value.length > 0) {
		if(! isNumber(domainform.dape.value)) {
			alert("������ϵ�˵绰�ֻ����벻�Ϸ���");
			domainform.dape.focus();
			return false;
		}
	}
	if(! isNumber(domainform.dafi.value)) {
		alert("������ϵ�˴���������Ų��Ϸ���");
		domainform.dafi.focus();
		return false;
	}
	if(! isNumber(domainform.dafa.value)) {
		alert("������ϵ�˴���������Ų��Ϸ���");
		domainform.dafa.focus();
		return false;
	}
	if(! isNumber(domainform.dafn.value)) {
		alert("������ϵ�˴�����벻�Ϸ���");
		domainform.dafn.focus();
		return false;
	}
	if(! isMail(domainform.daem.value)) {
		alert("������ϵ�˵����ʼ����Ϸ���");
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
<title>����-���ſƼ� ����ע�� A��¼��MX��¼��URLת���������޸ģ�����MyDNS���ܣ��������20��������������������޸�������Ϣ��������Ϊ����������ӵ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="keywords" content="www.dnyes.com-domain,����ע��,whois��ѯ,A��¼��IPָ��,MX��¼���ʼ���¼��,URLת������ַת����,ǿ�������MyDNS���ܣ��������20������������ע���˿�����������޸�������Ϣ!">
<meta name="description" content="www.dnyes.com-domain,����ע��,whois��ѯ,A��¼��IPָ��,MX��¼���ʼ���¼��,URLת������ַת����,ǿ�������MyDNS���ܣ��������20������������ע���˿�����������޸�������Ϣ!">
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
				
===<b>����ע�������</b>===
<table border="0" cellspacing="0" cellpadding="1" align="center" width="98%">
    <tr>
      <td>
        <table border="0" cellpadding="3" cellspacing="1" bgcolor="#000000">
<form class="form" method="POST" action="domainregsucc.asp" name="domainform" onsubmit="return doValidate()">
                              <tr bgcolor="#FFFFFF"> 
                                <td colspan="3"> ��д����ע������� 
                                </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">1. ����</td>
                                <td colspan="2"><font color=red><%=domain%><%=after%></td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%"><div align="left"></div></td>
                                <td colspan="2">&nbsp; </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">2. ע�ᵥλ</td>
                                <td colspan="2" >��</td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%"> 2.1 ��λ���ƣ������ģ�<br>
                                  �����ע�ᣬ��д�������� </td>
                                <td colspan="2">  
                                  <input name="donc" size="40" maxlength="60" value="<%=rs("namez")%>">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%"> 2.2 ��λ���ƣ���Ӣ�ģ�<br>
                                  �����ע�ᣬ��д�������� </td>
                                <td colspan="2">  
                                  <input name="done" size="40" maxlength="80" value="<%=rs("namee")%>">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">2.3 ��λ�����ˣ������ģ�</td>
                                <td colspan="2"> 
                                  <input name="doceoc" type="text" value="<%=rs("contactz")%>" size="20" maxlength="20">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">2.4 ��λ�����ˣ���Ӣ�ģ�<br>
                                  Ӣ�ĵ��պ���֮������пո���</td>
                                <td colspan="2"> 
                                  <input name="doceoe" type="text" value="<%=rs("contacte")%>" size="20" maxlength="30">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">2.5 ʡ�ݣ�</td>
                                <td colspan="2"> <select name="dos" size="1"> 
                                    <option value=<%=rs("shengz")%> selected><%=rs("shengz")%></option>
                                    <option value="Beijing">����</option>
                                    <option value="Hongkong">���</option>
                                    <option value="Aomen">����</option>
                                    <option value="Taiwan">̨��</option>
                                    <option value="Shanghai">�Ϻ�</option>
                                    <option value="Guangdong">�㶫</option>
                                    <option value="Shandong">ɽ��</option>
                                    <option value="Sichuan">�Ĵ�</option>
                                    <option value="Fujian">����</option>
                                    <option value="Jiangsu">����</option>
                                    <option value="Zhejiang">�㽭</option>
                                    <option value="Tianjin">���</option>
                                    <option value="Chongqing">����</option>
                                    <option value="Hebei">�ӱ�</option>
                                    <option value="Henan">����</option>
                                    <option value="Heilongjiang">������</option>
                                    <option value="Jinlin">����</option>
                                    <option value="Liaoning">����</option>
                                    <option value="Neimenggu">���ɹ�</option>
                                    <option value="Hainan">����</option>
                                    <option value="Shanxi">ɽ��</option>
                                    <option value="Shanxi3">����</option>
                                    <option value="Anhui">����</option>
                                    <option value="Jiangxi">����</option>
                                    <option value="Gansu">����</option>
                                    <option value="Xinjiang">�½�</option>
                                    <option value="Hubei">����</option>
                                    <option value="Hunan">����</option>
                                    <option value="Yunnan">����</option>
                                    <option value="Guangxi">����</option>
                                    <option value="Ningxia">����</option>
                                    <option value="Guizhou">����</option>
                                    <option value="Qinghai">�ຣ</option>
                                    <option value="Xizang">����</option>
                                  </select> </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">2.6 ���У������ģ�</td>
                                <td colspan="2">  
                                  <input maxlength="20" name="docc" size="20" value="<%=rs("cityz")%>">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">2.7 ���У���Ӣ�ģ�</td>
                                <td colspan="2">  
                                  <input maxlength="30" name="doce" size="20" value="<%=rs("citye")%>">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">2.8 �ֵ���ַ�������ģ�</td>
                                <td colspan="2">  
                                  <input maxlength="100" name="doac" size="40" value="<%=rs("dizhiz")%>">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">2.9 �ֵ���ַ����Ӣ�ģ�</td>
                                <td colspan="2">  
                                  <input maxlength="120" name="doae" size="40" value="<%=rs("dizhie")%>">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">2.10 �������룺</td>
                                <td colspan="2">  
                                  <input maxlength="8" name="doz" size="10" value="<%=rs("postnumber")%>">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%"><div align="left"></div></td>
                                <td colspan="2">&nbsp; </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%" rowspan="2">3.������ϵ��</td>
                                <td width="35%"> 
                                  <input type="radio" value="V1" name="R1" onClick="fill()">
                                  ʹ��ע�ᵥλ����Ϣ </td>
                                <td width="34%"> 
                                  <input type="radio" value="V2" name="R1" checked>
                                  ��д������Ϣ</td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td colspan="2"> <font color="#FF0000">���Ҫע������������뱣�ֹ�����ϵ�˵���Ϣ��ע�ᵥλ��Ϣһ�£�����ע�������������ʧ�ܡ� 
                                </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">3.1 �����������ģ�</td>
                                <td colspan="2">  
                                  <input name="danc" size="20" maxlength="20">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%"> 3.2 ��������Ӣ�ģ�<br>
                                  Ӣ�ĵ��պ���֮������пո��� </td>
                                <td colspan="2">  
                                  <input name="dane" size="20" maxlength="30">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">3.3 ��λ���ƣ������ģ�</td>
                                <td colspan="2">  
                                  <input maxlength="80" name="daoc" size="40">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">3.4 ��λ���ƣ���Ӣ�ģ�</td>
                                <td colspan="2">  
                                  <input maxlength="100" name="daoe" size="40">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">3.5 ʡ�ݣ�</td>
                                <td colspan="2">  
                                  <select name="das" size="1">
                                    <option value=<%=rs("shengz")%> selected><%=rs("shengz")%></option>
                                    <option value="Beijing">����</option>
                                    <option value="Hongkong">���</option>
                                    <option value="Aomen">����</option>
                                    <option value="Taiwan">̨��</option>
                                    <option value="Shanghai">�Ϻ�</option>
                                    <option value="Guangdong">�㶫</option>
                                    <option value="Shandong">ɽ��</option>
                                    <option value="Sichuan">�Ĵ�</option>
                                    <option value="Fujian">����</option>
                                    <option value="Jiangsu">����</option>
                                    <option value="Zhejiang">�㽭</option>
                                    <option value="Tianjin">���</option>
                                    <option value="Chongqing">����</option>
                                    <option value="Hebei">�ӱ�</option>
                                    <option value="Henan">����</option>
                                    <option value="Heilongjiang">������</option>
                                    <option value="Jinlin">����</option>
                                    <option value="Liaoning">����</option>
                                    <option value="Neimenggu">���ɹ�</option>
                                    <option value="Hainan">����</option>
                                    <option value="Shanxi">ɽ��</option>
                                    <option value="Shanxi3">����</option>
                                    <option value="Anhui">����</option>
                                    <option value="Jiangxi">����</option>
                                    <option value="Gansu">����</option>
                                    <option value="Xinjiang">�½�</option>
                                    <option value="Hubei">����</option>
                                    <option value="Hunan">����</option>
                                    <option value="Yunnan">����</option>
                                    <option value="Guangxi">����</option>
                                    <option value="Ningxia">����</option>
                                    <option value="Guizhou">����</option>
                                    <option value="Qinghai">�ຣ</option>
                                    <option value="Xizang">����</option>
                                  </select>
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">3.6 ���У������ģ�</td>
                                <td colspan="2">  
                                  <input maxlength="20" name="dacc" size="20">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">3.7 ���У���Ӣ�ģ�</td>
                                <td colspan="2">  
                                  <input maxlength="30" name="dace" size="20">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">3.8 �ֵ���ַ�������ģ�</td>
                                <td colspan="2">  
                                  <input maxlength="100" name="daac" size="40">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">3.9 �ֵ���ַ����Ӣ�ģ�</td>
                                <td colspan="2">  
                                  <input maxlength="120" name="daae" size="40">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">3.10 �������룺</td>
                                <td colspan="2">  
                                  <input maxlength="8" name="daz" size="10">
                                  </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">3.11 �绰��</td>
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
                                  �������� - �������� -&nbsp;&nbsp;&nbsp;&nbsp; �绰����&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- 
                                  �ֻ����� </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">3.12 ���棨��û�п���绰���룩��</td>
                                <td colspan="2">  
                                  <input name="dafi" size="5" maxlength="7" value="86">
                                  <b>-</b> 
                                  <input name="dafa" size="7" maxlength="7">
                                  <b>-</b> 
                                  <input name="dafn" size="15" maxlength="15">
                                  <input type="hidden" maxlength="7" name="dafe" size="7">
                                  <br>
                                  �������� - �������� -&nbsp;&nbsp;&nbsp;&nbsp; �������</td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td width="31%">3.13 �����ʼ���</td>
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
                                    ���������� ��</p>
                                  <p>��������������û�����ÿ��Բ��</p></td>
                                <td colspan="2">&nbsp; &nbsp; 
                                   
                                  <input type="radio" name="radiobutton"  checked  value="dns3"  onClick='rable("sdns")' style="border-style: none">
                                  ʹ��Ĭ��DNS������</td>
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
                                  ����д��DNS������ע�������Ч��DNS </td>
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
                                <td colspan="3">&nbsp; �ر���������Ϊ��ʹ������ݵ�ע���������˴�ֻ����дע������Ϣ������������ϵ�ˡ�������ϵ�ˡ�������ϵ����Ϣ������ע��ɹ����½��Ա����ϵͳ��Ѹ��ļ��ɡ�</td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td colspan="3" align="center">  
                                  <input class=botton type="submit" name="B1" value="��д���ˣ�����һ��">
                                  <input class=botton type="reset" value="ȫ�������������д" name="B2">
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

<!--FORM����-->
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