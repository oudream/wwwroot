<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<!--#include file="css.asp"-->
<%on error resume next%>
<script language="JavaScript">

function checkuid() {
  var letter =  "_0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"

	if (document.ADDUser.uid.value.length == 0) {
		alert("�������û��ʺ�.");
		document.ADDUser.uid.focus();
		return false;
	}
	for(i = 0; i < ADDUser.uid.value.length; i++) {
    if(letter.indexOf(ADDUser.uid.value.charAt(i)) == -1)
  {
  	  	alert("��Ա�ʺű���������ĸ���»�����ɣ�" );
    	ADDUser.uid.focus();
    	return false;
  }
	}
		return true;
}

function testusername() {
	if (checkuid()){
		username=ADDUser.uid.value;
		window.open('usertest.asp?uid='+username,'testusername','toolbar=no,location=no,drectores=no,status=no,menubar=no,scrollbars=no,resizable=no,width=300,height=200 ,top=100,left=200');
		
}

}

function getTECHSP()
			{
			if (ADDUser.TECHCC.value=="CHINA"){
				 td8.style.display="block";
				 td7.style.display="none"; 
				 td10.style.display="block";
				 td9.style.display="none"; }
			else
				 {td7.style.display="block";
				 td8.style.display="none"; 
				 td9.style.display="block";
				 td10.style.display="none"; }
			}

function fill_contact()
{
  ADDUser.contact_cn.value = ADDUser.com_cn.value;
  ADDUser.contact_en.value = ADDUser.com_en.value;
}

function unfill_contact()
{
  ADDUser.contact_cn.value = "";
  ADDUser.contact_en.value = "";
}

			
function CheckForm()
{
  var letter =  "_0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"

	if (document.ADDUser.uid.value.length == 0) {
		alert("�������û��ʺ�.");
		document.ADDUser.uid.focus();
		return false;
	}
	for(i = 0; i < ADDUser.uid.value.length; i++) {
    if(letter.indexOf(ADDUser.uid.value.charAt(i)) == -1)
  {
  	  	alert("��Ա�ʺű���������ĸ���»�����ɣ�" );
    	ADDUser.uid.focus();
    	return false;
  }
	}
	if (document.ADDUser.pwd.value.length == 0) {
		alert("����������.");
		document.ADDUser.pwd.focus();
		return false;
	}
		if (document.ADDUser.PW_Again.value.length == 0) {
		alert("��ȷ������.");
		document.ADDUser.PW_Again.focus();
		return false;
	}
	if (document.ADDUser.pwd.value != document.ADDUser.PW_Again.value) {
		alert("��������������벻һ��������������.");
		document.ADDUser.pwd.focus();
		return false;
	}if (document.ADDUser.com_cn.value.length == 0) {
		alert("��������ҵ�������������.");
		document.ADDUser.com_cn.focus();
		return false;
	}if (document.ADDUser.TECHCC.value=="CHINA"){
		if (document.ADDUser.TECHSP_CN.value.length == 0) {
			alert("��ѡ��ʡ�ݵ���������.");
			document.ADDUser.TECHSP_CN.focus();
			return false;}
	}else{
		if (document.ADDUser.TECHSP.value.length == 0) {
			alert("������ʡ�ݵ���������.");
			document.ADDUser.TECHSP.focus();
			return false;}
	}if (document.ADDUser.cityz.value.length == 0) {
		alert("��������е���������.");
		document.ADDUser.cityz.focus();
		return false;
	}if (document.ADDUser.addr.value.length == 0) {
		alert("�������ַ����������.");
		document.ADDUser.addr.focus();
		return false;
	}if (document.ADDUser.postnumber.value.length == 0) {
		alert("��������ϵ�˵���������.");
		document.ADDUser.postnumber.focus();
		return false;
	}if (document.ADDUser.tel1_area_code.value.length == 0) {
		alert("��������ϵ�绰�����ţ��Ա����ǿ���Ϊ������.");
		document.ADDUser.tel1_area_code.focus();
		return false;
	}if (document.ADDUser.tel1.value.length == 0) {
		alert("��������ϵ�绰���Ա����ǿ���Ϊ������.");
		document.ADDUser.tel1.focus();
		return false;
	}if (document.ADDUser.fax_area_code.value.length == 0) {
		alert("�����봫��绰�����ţ��Ա����ǿ���Ϊ������.");
		document.ADDUser.fax_area_code.focus();
		return false;
	}if (document.ADDUser.fax.value.length == 0) {
		alert("�����봫��绰���Ա����ǿ���Ϊ������.");
		document.ADDUser.fax.focus();
		return false;
	}if (document.ADDUser.email.value.length == 0) {
		alert("������Email.");
		document.ADDUser.email.focus();
		return false;
	}if (document.ADDUser.email.value.length > 0 && !document.ADDUser.email.value.match( /^.+@.+$/ ) ) {
		alert("Email ��������������");
		document.ADDUser.email.focus();
		return false;
	}
	
	
   if(ADDUser.uid.value.length < 4)
  {
    alert("�ʺŲ�������6���ַ���")
    ADDUser.uid.focus()
    return false
  }
  if(ADDUser.uid.value.length > 20)
  {
    alert("�ʺŲ��ܳ���20���ַ���")
    ADDUser.uid.focus()
    return false
  }      
   if(ADDUser.pwd.value.length < 6)
  {
    alert("���벻������6���ַ���")
    ADDUser.pwd.focus()
    return false
  }
  if(ADDUser.pwd.value.length > 20)
  {
    alert("���벻�ܳ���20���ַ���")
    ADDUser.pwd.focus()
    return false
  }      
  
  
   if(isNaN(ADDUser.postnumber.value))
  {
    alert("������������������д")
    ADDUser.postnumber.focus()
    return false
  }
   if(isNaN(ADDUser.tel1_gov_code.value))
  {
    alert("������������������д")
    ADDUser.tel1_gov_code.focus()
    return false
  }
   if(isNaN(ADDUser.tel1_area_code.value))
  {
    alert("��������������д")
    ADDUser.tel1_area_code.focus()
    return false
  }
   if(isNaN(ADDUser.tel1.value))
  {
    alert("�绰��������������д")
    ADDUser.tel1.focus()
    return false
  }
   if(isNaN(ADDUser.tel1_ext_code.value))
  {
    alert("�绰�ֻ���������������д")
    ADDUser.tel1_ext_code.focus()
    return false
  }
   if(isNaN(ADDUser.tel2_gov_code.value))
  {
    alert("������������������д")
    ADDUser.tel2_gov_code.focus()
    return false
  }
   if(isNaN(ADDUser.tel2_area_code.value))
  {
    alert("��������������д")
    ADDUser.tel2_area_code.focus()
    return false
  }
   if(isNaN(ADDUser.tel2.value))
  {
    alert("�绰��������������д")
    ADDUser.tel2.focus()
    return false
  }
   if(isNaN(ADDUser.tel2_ext_code.value))
  {
    alert("�绰�ֻ���������������д")
    ADDUser.tel2_ext_code.focus()
    return false
  }
   if(isNaN(ADDUser.fax_gov_code.value))
  {
    alert("���������������������д")
    ADDUser.fax_gov_code.focus()
    return false
  }
   if(isNaN(ADDUser.fax_area_code.value))
  {
    alert("������������������д")
    ADDUser.fax_area_code.focus()
    return false
  }
   if(isNaN(ADDUser.fax.value))
  {
    alert("����绰��������������д")
    ADDUser.fax.focus()
    return false
  }
   if(isNaN(ADDUser.fax_ext_code.value))
  {
    alert("����绰�ֻ���������������д")
    ADDUser.fax_ext_code.focus()
    return false
  }

	
	
	
	return true;
}
</script>
<html>
<head>
<title>�û�ע��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="dnyes/inc/southdns.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<!--#include file="top.asp"-->
<%
if session("y_it_uid")<> "" then
response.Redirect("error.asp?err=013")
response.End()
end if
%>
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
          <td width="229" valign="top" background="images/left-228x5.gif"> <table width="229" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="229" height="53" background="images/left-1b.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="10%" height="18" valign="bottom">&nbsp; </td>
                      <td width="8%" valign="bottom"><img src="images/gogo.gif" width="6" height="15"></td>
                      <td width="82%" valign="bottom"><font color="#FFFFFF"><b>�� 
                        �� �� ¼</b></font></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td valign="top"> <table width="228" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td background="images/left-1-left-2b.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td width="21" valign="top"><img src="images/left-1-left-1b.gif" width="21" height="190"></td>
                            <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                                <tr> 
                                  <td> 
                                    <!--#include file="userlogin.asp"-->
                                  </td>
                                </tr>
                                <tr> 
                                  <td>
                                    <!--#include file="userhelp.asp" -->
                                  </td>
                                </tr>
                              </table></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td><img src="images/left-1-left-3b.gif" width="229" height="12"></td>
                    </tr>
                  </table></td>
              </tr>
            </table>
            <table width="229" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="229" height="28" background="images/left-2.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="10%" height="18" valign="bottom">&nbsp; </td>
                      <td width="8%" valign="bottom"><img src="images/gogo.gif" width="6" height="15"></td>
                      <td width="82%" valign="bottom"><font color="#FFFFFF"><b>�� 
                        �� �� ��</b></font></td>
                    </tr>
                  </table></td>
              </tr>
            </table>
            <TABLE width=70% border=0 cellPadding=0 cellSpacing=0>
              <TBODY>
                <TR> 
                  <TD height="3" 
                vAlign=top 
                style="BORDER-left: #000000 1px solid"><img src="1.gif" width="1" height="1"></TD>
                </TR>
              </TBODY>
            </TABLE>
            <table width="228" height="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="1" background="images/1x2-black.gif"><img src="images/1x2-black.gif" width="1" height="2"></td>
                <td height="180" valign="top"> 
                  <!--#include file="support.asp" -->
                </td>
              </tr>
            </table></td>
          <td width="7">&nbsp;</td>
          <td width="513" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr> 
                <td colspan="5" height="88"> <TABLE cellSpacing=0 cellPadding=0 width=100% border=0>
                    <TBODY>
                      <TR> 
                        <TD width="34%" vAlign=center style="BORDER-left: #CCCCCC 1px solid; BORDER-RIGHT: #CCCCCC 1px solid; BORDER-BOTTOM: #CCCCCC 1px solid;BORDER-TOP: #CCCCCC 1px solid"> 
                          <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td height="8" colspan="2"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td width="12" height="25" background="images/host-2.gif">&nbsp;</td>
                                    <td valign="bottom" background="images/host-2.gif">�� �� ע �� >> �� ҵ �� �� ע �� ��</td>
                                  </tr>
                                </table>
                                <table width="75%" border="0" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td height="6"><img src="1.gif" width="1" height="1"></td>
                                  </tr>
                                </table>
                                
                                <table width="498" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                                  <tr> 
                                    <td bgcolor="#f5f5f5"> ��ʾ��<br>
                                      <br>
                                      1.ע��Ϊ�û������������߶������ǵĲ�Ʒ�ͷ���<br>
                                      <br>
                                      2.����Զ������Ķ��������ߵĹ���<br>
                                      <br>
                                      3.�������������ʱ����Щ���Ͽ��԰���˳�����һ�����<br>
                                      <br>
                                      ���ԣ�<font color="#FF0033">����������д��������</font>. 
                                      ����<b><font color="#FF3333">*</font></b>Ϊ������д����Ŀ�� 
                                    </td>
                                  </tr>
                                </table>
                                <br>
                                <table width="498" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                                  <form name="ADDUser" method="POST" action="userregfix.asp" onSubmit="return CheckForm();">
                                    <tr> 
                                      <td width="158" bgcolor="#f5f5f5">�û��ʺţ�</td>
                                      <td colspan="2" bgcolor="#FFFFFF"> <input type="text" name="uid" maxlength="21" class="form" size="21">
									  &nbsp;--&nbsp; 
										 <INPUT class="botton" name=testname onclick=javascript:testusername() type=button value=�û��ʺŲ���>
                                      </td>
                                      <td width="50" colspan="3" bgcolor="#FFFFFF" ><font color="#CC0000">*</font></td>
                                    </tr>
                                    <tr> 
                                      <td bgcolor="#f5f5f5">�û����룺</td>
                                      <td colspan="2" bgcolor="#FFFFFF"><input type="password" name="pwd" maxlength="21" class="form" size="21"></td>
                                      <td colspan="3" bgcolor="#FFFFFF" ><font color="#CC0000">*</font></td>
                                    </tr>
                                    <tr> 
                                      <td bgcolor="#f5f5f5">�ظ����룺</td>
                                      <td colspan="2" bgcolor="#FFFFFF"><input type="password" name="PW_Again" maxlength="21" class="form" size="21"></td>
                                      <td colspan="3" bgcolor="#FFFFFF" ><font color="#CC0000">*</font></td>
                                    </tr>
                                    <tr> 
                                      <td bgcolor="#f5f5f5">��ҵ���ƣ����ģ���</td>
                                      <td colspan="2" bgcolor="#FFFFFF"><input name="com_cn" type="text" size="50"  maxlength="80" ></td>
                                      <td colspan="3" bgcolor="#FFFFFF" ><font color="#CC0000">*</font></td>
                                    </tr>
                                    <tr> 
                                      <td bgcolor="#f5f5f5">��ҵ���ƣ�Ӣ�ģ���</td>
                                      <td colspan="2" bgcolor="#FFFFFF"><input name="com_en" type="text" size="50"  maxlength="120"></td>
                                      <td colspan="3" bgcolor="#FFFFFF" >&nbsp;</td>
                                    </tr>
                                    <tr> 
                                      <td height="36" bgcolor="#f5f5f5">��ϵ�ˣ����ģ���</td>
                                      <td width="188" colspan="2" bgcolor="#FFFFFF"><input name="contact_cn" type="text"  size="21" maxlength="21"></td>
                                      <td colspan="3" bgcolor="#FFFFFF" ><font color="#CC0000">*</font></td>
                                    </tr>
                                    <tr> 
                                      <td height="20" bgcolor="#f5f5f5">��ϵ�ˣ�Ӣ�ģ���</td>
                                      <td width="188" colspan="2" bgcolor="#FFFFFF"><input name="contact_en" type="text"  size="21" maxlength="30"></td>
                                      <td colspan="3" bgcolor="#FFFFFF" >&nbsp;</td>
                                    </tr>
                                    <tr> 
                                      <td bgcolor="#f5f5f5">��ϵ���Ա�</td>
                                      <td colspan="2" bgcolor="#FFFFFF"><input name="u_or_i" type="radio" value="1"  checked>
                                        �� 
                                        <input name="u_or_i" type="radio" value="0" >
                                        Ů</td>
                                      <td colspan="3" bgcolor="#FFFFFF" ><font color="#CC0000">*</font></td>
                                    </tr>
                                    <tr> 
                                      <td bgcolor="#f5f5f5">���һ������</td>
                                      <td colspan="2" bgcolor="#FFFFFF"><select name="TECHCC"  onChange="getTECHSP()">
                                          <option value="CHINA" selected>�й� CHINA</option>
                                          <option value="AFGHANISTAN">AFGHANISTAN</option>
                                          <option value="ALBANIA">ALBANIA</option>
                                          <option value="ALGERIA">ALGERIA</option>
                                          <option value="AMERICAN SAMOA">AMERICAN 
                                          SAMOA</option>
                                          <option value="ANDORRA">ANDORRA</option>
                                          <option value="ANGOLA">ANGOLA</option>
                                          <option value="ANGUILLA">ANGUILLA</option>
                                          <option value="ANTARCTICA">ANTARCTICA</option>
                                          <option value="ANTIGUA AND BARBUDA">ANTIGUA 
                                          AND BARBUDA</option>
                                          <option value="ARGENTINA">ARGENTINA</option>
                                          <option value="ARMENIA">ARMENIA</option>
                                          <option value="ARUBA">ARUBA</option>
                                          <option value="AUSTRALIA">AUSTRALIA</option>
                                          <option value="AUSTRI">AUSTRI</option>
                                          <option value="AZERBAIJAN">AZERBAIJAN</option>
                                          <option value="BAHAMAS">BAHAMAS</option>
                                          <option value="BAHRAIN">BAHRAIN</option>
                                          <option value="BANGLADESH">BANGLADESH</option>
                                          <option value="BARBADOS">BARBADOS</option>
                                          <option value="BELARUS">BELARUS</option>
                                          <option value="BELGIUM">BELGIUM</option>
                                          <option value="BELIZE">BELIZE</option>
                                          <option value="BENIN">BENIN</option>
                                          <option value="BERMUDA">BERMUDA</option>
                                          <option value="BHUTAN">BHUTAN</option>
                                          <option value="BOLIVIA">BOLIVIA</option>
                                          <option value="BOSNIA AND HERZEGOVINA">BOSNIA 
                                          AND HERZEGOVINA</option>
                                          <option value="BOTSWANA">BOTSWANA</option>
                                          <option value="BOUVET ISLAND">BOUVET 
                                          ISLAND</option>
                                          <option value="BRAZIL">BRAZIL</option>
                                          <option value="BRITISH INDIAN OCEAN TERRITORY">BRITISH 
                                          INDIAN OCEAN TERRITORY</option>
                                          <option value="BRUNEI DARUSSALAM">BRUNEI 
                                          DARUSSALAM</option>
                                          <option value="BULGARIA">BULGARIA</option>
                                          <option value="BURKINA FASO">BURKINA 
                                          FASO</option>
                                          <option value="BURUNDI">BURUNDI</option>
                                          <option value="CAMBODIA">CAMBODIA</option>
                                          <option value="CAMEROON">CAMEROON</option>
                                          <option value="CANADA">CANADA</option>
                                          <option value="CAPE VERDE">CAPE VERDE</option>
                                          <option value="CAYMAN ISLANDS">CAYMAN 
                                          ISLANDS</option>
                                          <option value="CENTRAL AFRICAN REPUBLIC">CENTRAL 
                                          AFRICAN REPUBLIC</option>
                                          <option value="CHAD">CHAD</option>
                                          <option value="CHILE">CHILE</option>
                                          <option value="CHRISTMAS ISLAND">CHRISTMAS 
                                          ISLAND</option>
                                          <option value="COCOS (KEELING) ISLANDS">COCOS 
                                          (KEELING) ISLANDS</option>
                                          <option value="COLOMBIA">COLOMBIA</option>
                                          <option value="COMOROS">COMOROS</option>
                                          <option value="CONGO">CONGO</option>
                                          <option value="CONGO, THE DEMOCRATIC REPUBLIC OF THE">CONGO, 
                                          THE DEMOCRATIC REPUBLIC OF THE</option>
                                          <option value="COOK ISLANDS">COOK ISLANDS</option>
                                          <option value="COSTA RICA">COSTA RICA</option>
                                          <option value="C?TE D' IVOIR">C?TE D' 
                                          IVOIRE</option>
                                          <option value="CROATIA">CROATIA</option>
                                          <option value="CUBA">CUBA</option>
                                          <option value="CYPRUS">CYPRUS</option>
                                          <option value="CZECH REPUBLIC">CZECH 
                                          REPUBLIC</option>
                                          <option value="DENMARK">DENMARK</option>
                                          <option value="DJIBOUTI">DJIBOUTI</option>
                                          <option value="DOMINICA">DOMINICA</option>
                                          <option value="DOMINICAN REPUBLIC">DOMINICAN 
                                          REPUBLIC</option>
                                          <option value="EAST TIMOR">EAST TIMOR</option>
                                          <option value="ECUADOR">ECUADOR</option>
                                          <option value="EGYPT">EGYPT</option>
                                          <option value="EL SALVADOR">EL SALVADOR</option>
                                          <option value="EQUATORIAL GUINEA">EQUATORIAL 
                                          GUINEA</option>
                                          <option value="ERITREA">ERITREA</option>
                                          <option value="ESTONIA">ESTONIA</option>
                                          <option value="ETHIOPIA">ETHIOPIA</option>
                                          <option value="FALKLAND ISLANDS (MALVINAS)">FALKLAND 
                                          ISLANDS (MALVINAS)</option>
                                          <option value="FAROE ISLANDS">FAROE 
                                          ISLANDS</option>
                                          <option value="FIJI">FIJI</option>
                                          <option value="FINLAND">FINLAND</option>
                                          <option value="FRANCE">FRANCE</option>
                                          <option value="FRENCH GUIANA">FRENCH 
                                          GUIANA</option>
                                          <option value="FRENCH POLYNESIA">FRENCH 
                                          POLYNESIA</option>
                                          <option value="FRENCH SOUTHERN TERRITORIES">FRENCH 
                                          SOUTHERN TERRITORIES</option>
                                          <option value="GABON">GABON</option>
                                          <option value="GAMBIA">GAMBIA</option>
                                          <option value="GEORGIA">GEORGIA</option>
                                          <option value="GERMANY">GERMANY</option>
                                          <option value="GHANA">GHANA</option>
                                          <option value="GIBRALTAR">GIBRALTAR</option>
                                          <option value="GREECE">GREECE</option>
                                          <option value="GREENLAND">GREENLAND</option>
                                          <option value="GRENADA">GRENADA</option>
                                          <option value="GUADELOUPE">GUADELOUPE</option>
                                          <option value="GUAM">GUAM</option>
                                          <option value="GUATEMALA">GUATEMALA</option>
                                          <option value="GUINEA">GUINEA</option>
                                          <option value="GUINEA-BISSAU">GUINEA-BISSAU</option>
                                          <option value="GUYANA">GUYANA</option>
                                          <option value="HAITI">HAITI</option>
                                          <option value="HEARD ISLAND AND MCDONALD ISLANDS">HEARD 
                                          ISLAND AND MCDONALD ISLANDS</option>
                                          <option value="HOLY SEE (VATICAN CITY STATE)">HOLY 
                                          SEE (VATICAN CITY STATE)</option>
                                          <option value="HONDURAS">HONDURAS</option>
                                          <option value="HONG KONG">HONG KONG</option>
                                          <option value="HUNGARY">HUNGARY</option>
                                          <option value="ICELAND">ICELAND</option>
                                          <option value="INDIA">INDIA</option>
                                          <option value="INDONESIA">INDONESIA</option>
                                          <option value="IRAN, ISLAMIC REPUBLIC OF">IRAN, 
                                          ISLAMIC REPUBLIC OF</option>
                                          <option value="IRAQ">IRAQ</option>
                                          <option value="IRELAND">IRELAND</option>
                                          <option value="ISRAEL">ISRAEL</option>
                                          <option value="ITALY">ITALY</option>
                                          <option value="JAMAICA">JAMAICA</option>
                                          <option value="JAPAN">JAPAN</option>
                                          <option value="JORDAN">JORDAN</option>
                                          <option value="KAZAKSTAN">KAZAKSTAN</option>
                                          <option value="KENYA">KENYA</option>
                                          <option value="KIRIBATI">KIRIBATI</option>
                                          <option value="KOREA, DEMOCRATIC PEOPLE'S REPUBLIC OF">KOREA, 
                                          DEMOCRATIC PEOPLE'S REPUBLIC OF</option>
                                          <option value="KOREA, REPUBLIC OF">KOREA, 
                                          REPUBLIC OF</option>
                                          <option value="KUWAIT">KUWAIT</option>
                                          <option value="KYRGYZSTAN">KYRGYZSTAN</option>
                                          <option value="LAO PEOPLE'S DEMOCRATIC REPUBLIC<">LAO 
                                          PEOPLE'S DEMOCRATIC REPUBLIC</option>
                                          <option value="LATVIA">LATVIA</option>
                                          <option value="LEBANON">LEBANON</option>
                                          <option value="LESOTHO">LESOTHO</option>
                                          <option value="LIBERIA">LIBERIA</option>
                                          <option value="LIBYAN ARAB JAMAHIRIYA">LIBYAN 
                                          ARAB JAMAHIRIYA</option>
                                          <option value="LIECHTENSTEIN">LIECHTENSTEIN</option>
                                          <option value="LITHUANIA">LITHUANIA</option>
                                          <option value="LUXEMBOURG">LUXEMBOURG</option>
                                          <option value="MACAU">MACAU</option>
                                          <option value="MACEDONIA, THE FORMER YUGOSLAV REPUBLIC OF">MACEDONIA, 
                                          THE FORMER YUGOSLAV REPUBLIC OF</option>
                                          <option value="MADAGASCAR">MADAGASCAR</option>
                                          <option value="MALAWI">MALAWI</option>
                                          <option value="MALAYSIA">MALAYSIA</option>
                                          <option value="MALDIVES">MALDIVES</option>
                                          <option value="MALI">MALI</option>
                                          <option value="MALTA">MALTA</option>
                                          <option value="MARSHALL ISLANDS">MARSHALL 
                                          ISLANDS</option>
                                          <option value="MARTINIQUE">MARTINIQUE</option>
                                          <option value="MAURITANIA">MAURITANIA</option>
                                          <option value="MAURITIUS">MAURITIUS</option>
                                          <option value="MAYOTTE">MAYOTTE</option>
                                          <option value="MEXICO">MEXICO</option>
                                          <option value="MICRONESIA, FEDERATED STATES OF">MICRONESIA, 
                                          FEDERATED STATES OF</option>
                                          <option value="MOLDOVA, REPUBLIC OF">MOLDOVA, 
                                          REPUBLIC OF</option>
                                          <option value="MONACO">MONACO</option>
                                          <option value="MONGOLIA">MONGOLIA</option>
                                          <option value="MONTSERRAT">MONTSERRAT</option>
                                          <option value="MOROCCO">MOROCCO</option>
                                          <option value="MOZAMBIQUE">MOZAMBIQUE</option>
                                          <option value="MYANMAR">MYANMAR</option>
                                          <option value="NAMIBIA">NAMIBIA</option>
                                          <option value="NAURU">NAURU</option>
                                          <option value="NEPAL">NEPAL</option>
                                          <option value="NETHERLANDS">NETHERLANDS</option>
                                          <option value="NETHERLANDS ANTILLES">NETHERLANDS 
                                          ANTILLES</option>
                                          <option value="NEW CALEDONIA">NEW CALEDONIA</option>
                                          <option value="NEW ZEALAND">NEW ZEALAND</option>
                                          <option value="NICARAGUA">NICARAGUA</option>
                                          <option value="NIGER">NIGER</option>
                                          <option value="NIGERIA">NIGERIA</option>
                                          <option value="NIUE">NIUE</option>
                                          <option value="NORFOLK ISLAND">NORFOLK 
                                          ISLAND</option>
                                          <option value="NORTHERN MARIANA ISLANDS">NORTHERN 
                                          MARIANA ISLANDS</option>
                                          <option value="NORWAY">NORWAY</option>
                                          <option value="OMAN">OMAN</option>
                                          <option value="PAKISTAN">PAKISTAN</option>
                                          <option value="PALAU">PALAU</option>
                                          <option value="PALESTINIAN TERRITORY, OCCUPIED">PALESTINIAN 
                                          TERRITORY, OCCUPIED</option>
                                          <option value="PANAMA">PANAMA</option>
                                          <option value="PAPUA NEW GUINEA">PAPUA 
                                          NEW GUINEA</option>
                                          <option value="PARAGUAY">PARAGUAY</option>
                                          <option value="PERU">PERU</option>
                                          <option value="PHILIPPINES">PHILIPPINES</option>
                                          <option value="PITCAIRN">PITCAIRN</option>
                                          <option value="POLAND">POLAND</option>
                                          <option value="PORTUGAL">PORTUGAL</option>
                                          <option value="PUERTO RICO">PUERTO RICO</option>
                                          <option value="QATAR">QATAR</option>
                                          <option value="R��UNION">R��UNION</option>
                                          <option value="ROMANIA">ROMANIA</option>
                                          <option value="RUSSIAN FEDERATION">RUSSIAN 
                                          FEDERATION</option>
                                          <option value="RWANDA">RWANDA</option>
                                          <option value="SAINT HELENA">SAINT HELENA</option>
                                          <option value="SAINT KITTS AND NEVIS">SAINT 
                                          KITTS AND NEVIS</option>
                                          <option value="SAINT LUCIA">SAINT LUCIA</option>
                                          <option value="SAINT PIERRE AND MIQUELON">SAINT 
                                          PIERRE AND MIQUELON</option>
                                          <option value="SAINT VINCENT AND THE GRENADINES">SAINT 
                                          VINCENT AND THE GRENADINES</option>
                                          <option value="SAMOA">SAMOA</option>
                                          <option value="SAN MARINO">SAN MARINO</option>
                                          <option value="SAO TOME AND PRINCIPE">SAO 
                                          TOME AND PRINCIPE</option>
                                          <option value="SAUDI ARABIA">SAUDI ARABIA</option>
                                          <option value="SENEGAL">SENEGAL</option>
                                          <option value="SEYCHELLES">SEYCHELLES</option>
                                          <option value="SIERRA LEONE">SIERRA 
                                          LEONE</option>
                                          <option value="SINGAPORE">SINGAPORE</option>
                                          <option value="SLOVAKIA">SLOVAKIA</option>
                                          <option value="SLOVENIA">SLOVENIA</option>
                                          <option value="SOLOMON ISLANDS">SOLOMON 
                                          ISLANDS</option>
                                          <option value="SOMALIA">SOMALIA</option>
                                          <option value="SOUTH AFRICA">SOUTH AFRICA</option>
                                          <option value="SOUTH GEORGIA AND THE SOUTH SANDWICH ISLANDS">SOUTH 
                                          GEORGIA AND THE SOUTH SANDWICH ISLANDS</option>
                                          <option value="SPAIN">SPAIN</option>
                                          <option value="SRI LANKA">SRI LANKA</option>
                                          <option value="SUDAN">SUDAN</option>
                                          <option value="SURINAME">SURINAME</option>
                                          <option value="SVALBARD AND JAN MAYEN">SVALBARD 
                                          AND JAN MAYEN</option>
                                          <option value="SWAZILAND">SWAZILAND</option>
                                          <option value="SWEDEN">SWEDEN</option>
                                          <option value="SWITZERLAND">SWITZERLAND</option>
                                          <option value="SYRIAN ARAB REPUBLIC">SYRIAN 
                                          ARAB REPUBLIC</option>
                                          <option value="CHINAESE TAIBEI">CHINAESE 
                                          TAIBEI</option>
                                          <option value="TAJIKISTAN">TAJIKISTAN</option>
                                          <option value="TANZANIA, UNITED REPUBLIC OF">TANZANIA, 
                                          UNITED REPUBLIC OF</option>
                                          <option value="THAILAND">THAILAND</option>
                                          <option value="TOGO">TOGO</option>
                                          <option value="TOKELAU">TOKELAU</option>
                                          <option value="TONGA">TONGA</option>
                                          <option value="TRINIDAD AND TOBAGO">TRINIDAD 
                                          AND TOBAGO</option>
                                          <option value="TUNISIA">TUNISIA</option>
                                          <option value="TURKEY">TURKEY</option>
                                          <option value="TURKMENISTAN">TURKMENISTAN</option>
                                          <option value="TURKS AND CAICOS ISLANDS">TURKS 
                                          AND CAICOS ISLANDS</option>
                                          <option value="TUVALU">TUVALU</option>
                                          <option value="UGANDA">UGANDA</option>
                                          <option value="UKRAINE">UKRAINE</option>
                                          <option value="UNITED ARAB EMIRATES">UNITED 
                                          ARAB EMIRATES</option>
                                          <option value="UNITED KINGDOM">UNITED 
                                          KINGDOM</option>
                                          <option value="UNITED STATES">UNITED 
                                          STATES</option>
                                          <option value="UNITED STATES MINOR OUTLYING ISLANDS">UNITED 
                                          STATES MINOR OUTLYING ISLANDS</option>
                                          <option value="URUGUAY">URUGUAY</option>
                                          <option value="UZBEKISTAN">UZBEKISTAN</option>
                                          <option value="VANUATU">VANUATU</option>
                                          <option value="VENEZUELA">VENEZUELA</option>
                                          <option value="VIET NAM">VIET NAM</option>
                                          <option value="VIRGIN ISLANDS, BRITISH">VIRGIN 
                                          ISLANDS, BRITISH</option>
                                          <option value="VIRGIN ISLANDS, U.S.">VIRGIN 
                                          ISLANDS, U.S.</option>
                                          <option value="WALLIS AND FUTUNA">WALLIS 
                                          AND FUTUNA</option>
                                          <option value="WESTERN SAHARA">WESTERN 
                                          SAHARA</option>
                                          <option value="YEMEN">YEMEN</option>
                                          <option value="YUGOSLAVIA">YUGOSLAVIA</option>
                                          <option value="ZAMBIA">ZAMBIA</option>
                                          <option value="ZIMBABWE">ZIMBABWE</option>
                                        </select></td>
                                      <td colspan="3" bgcolor="#FFFFFF" ><font color="#CC0000">*</font></td>
                                    </tr>
                                    <tr> 
                                      <td bgcolor="#f5f5f5">ʡ��(����)��</td>
                                      <td colspan="2" bgcolor="#FFFFFF" id="td7" style="display:none;"><input name="TECHSP" type="text" value="" size="20" maxlength="30"></td>
                                      <td colspan="2" bgcolor="#FFFFFF" id="td8"><select name="TECHSP_CN" >
                                          <option value="" selected>��ѡ��...</option>
                                          <option value="����">����</option>
                                          <option value="�Ϻ�">�Ϻ�</option>
                                          <option value="���">���</option>
                                          <option value="����">����</option>
                                          <option value="���">���</option>
                                          <option value="�㶫">�㶫</option>
                                          <option value="ɽ��">ɽ��</option>
                                          <option value="ɽ��">ɽ��</option>
                                          <option value="�Ĵ�">�Ĵ�</option>
                                          <option value="����">����</option>
                                          <option value="����">����</option>
                                          <option value="�㽭">�㽭</option>
                                          <option value="�ӱ�">�ӱ�</option>
                                          <option value="����">����</option>
                                          <option value="������">������</option>
                                          <option value="����">����</option>
                                          <option value="����">����</option>
                                          <option value="���ɹ�">���ɹ�</option>
                                          <option value="����">����</option>
                                          <option value="����">����</option>
                                          <option value="����">����</option>
                                          <option value="����">����</option>
                                          <option value="����">����</option>
                                          <option value="�½�">�½�</option>
                                          <option value="����">����</option>
                                          <option value="����">����</option>
                                          <option value="����">����</option>
                                          <option value="����">����</option>
                                          <option value="����">����</option>
                                          <option value="����">����</option>
                                          <option value="�ຣ">�ຣ</option>
                                          <option value="����">����</option>
                                          <option value="���">���</option>
                                          <option value="̨��">̨��</option>
                                          <option value="����">����</option>
                                        </select> </td>
                                      <td colspan="3" bgcolor="#FFFFFF" ><font color="#CC0000">*</font></td>
                                    </tr>
                                    <tr> 
                                      <td bgcolor="#f5f5f5">ʡ��(Ӣ��)��</td>
                                      <td colspan="2" bgcolor="#FFFFFF" id="td9" style="display:none;"> 
                                        <input name="TECHEN" type="text" value="" size="20" maxlength="50"> 
                                        </td>
                                      <td colspan="2" bgcolor="#FFFFFF" id="td10" ><font color="#000000"><span style="font-size: 9pt"> 
                                        <select name="TECHEN_CN">
                                          <option value="" selected>select...</option>
                                          <option value="Beijing">Beijing</option>
                                          <option value="Shanghai">Shanghai</option>
                                          <option value="Tianjin">Tianjin</option>
                                          <option value="Chongqing">Chongqing</option>
                                          <option value="Guangdong">Guangdong</option>
                                          <option value="Shandong">Shandong</option>
                                          <option value="Shanxi">Shanxi</option>
                                          <option value="Sichuan">Sichuan</option>
                                          <option value="Fujian">Fujian</option>
                                          <option value="Jiangsu">Jiangsu</option>
                                          <option value="Zhejiang">Zhejiang</option>
                                          <option value="Hebei">Hebei</option>
                                          <option value="Henan">Henan</option>
                                          <option value="Heilongjiang">Heilongjiang</option>
                                          <option value="Jinlin">Jinlin</option>
                                          <option value="Liaoning">Liaoning</option>
                                          <option value="Neimenggu">Neimenggu</option>
                                          <option value="Hainan">Hainan</option>
                                          <option value="Shanxi3">Shanxi3</option>
                                          <option value="Anhui">Anhui</option>
                                          <option value="Jiangxi">Jiangxi</option>
                                          <option value="Gansu">Gansu</option>
                                          <option value="Xinjiang">Xinjiang</option>
                                          <option value="Hubei">Hubei</option>
                                          <option value="Hunan">Hunan</option>
                                          <option value="Yunnan">Yunnan</option>
                                          <option value="Guangxi">Guangxi</option>
                                          <option value="Ningxia">Ningxia</option>
                                          <option value="Guizhou">Guizhou</option>
                                          <option value="Qinghai">Qinghai</option>
                                          <option value="Xizang">Xizang</option>
                                          <option value="Hongkong">Hongkong</option>
                                          <option value="Aomen">Aomen</option>
                                          <option value="Taiwan">Taiwan</option>
                                        </select>
                                        <font color="#CC0000"> </font> </span></font></td>
                                      <td colspan="3" bgcolor="#FFFFFF">&nbsp;</td>
                                    </tr>
                                    <tr> 
                                      <td bgcolor="#f5f5f5">���У����ģ���</td>
                                      <td colspan="2" bgcolor="#FFFFFF"><input type="text" name="cityz" size="20" maxlength="30" class="form"></td>
                                      <td colspan="3" bgcolor="#FFFFFF" ><font color="#CC0000">*</font></td>
                                    </tr>
                                    <tr> 
                                      <td bgcolor="#f5f5f5">���У�Ӣ�ģ��� </td>
                                      <td colspan="2" bgcolor="#FFFFFF"><input type="text" name="citye"  maxlength="50" size="20" class="form"></td>
                                      <td colspan="3" bgcolor="#FFFFFF" >&nbsp;</td>
                                    </tr>
                                    <tr> 
                                      <td bgcolor="#f5f5f5">ͨ�ŵ�ַ�����ģ���</td>
                                      <td colspan="2" bgcolor="#FFFFFF"><input type="text" name="addr" size="50" maxlength="80" class="form"></td>
                                      <td colspan="3" bgcolor="#FFFFFF" ><font color="#CC0000">*</font></td>
                                    </tr>
                                    <tr> 
                                      <td bgcolor="#f5f5f5">ͨ�ŵ�ַ��Ӣ�ģ���</td>
                                      <td colspan="2" bgcolor="#FFFFFF"><input type="text" name="addr_en"  maxlength="120" size="50" class="form"></td>
                                      <td colspan="3" bgcolor="#FFFFFF" >&nbsp;</td>
                                    </tr>
                                    <tr> 
                                      <td bgcolor="#f5f5f5">�������룺</td>
                                      <td colspan="2" bgcolor="#FFFFFF"><input type="text" name="postnumber" size="12" maxlength="10" class="form"></td>
                                      <td colspan="3" bgcolor="#FFFFFF" ><font color="#CC0000">*</font></td>
                                    </tr>
                                    <tr> 
                                      <td bgcolor="#f5f5f5">��ϵ�绰1�� </td>
                                      <td colspan="2" bgcolor="#FFFFFF"> <input type="text" name="tel1_gov_code"  maxlength="5" size="3" class="form" value="86">
                                        - 
                                        <input type="text" name="tel1_area_code"  maxlength="7" size="5" class="form">
                                        - 
                                        <input type="text" name="tel1"  maxlength="15" size="10" class="form">
                                        - 
                                        <input type="text" name="tel1_ext_code"  maxlength="7" size="5" class="form"> 
                                      </td>
                                      <td colspan="3" bgcolor="#FFFFFF" ><font color="#CC0000">*</font></td>
                                    </tr>
                                    <tr> 
                                      <td bgcolor="#f5f5f5">��ϵ�绰2�� </td>
                                      <td colspan="2" bgcolor="#FFFFFF"> <input type="text" name="tel2_gov_code"  maxlength="5" size="3" class="form" value="86">
                                        - 
                                        <input type="text" name="tel2_area_code"  maxlength="7" size="5" class="form">
                                        - 
                                        <input type="text" name="tel2"  maxlength="15" size="10" class="form">
                                        - 
                                        <input type="text" name="tel2_ext_code"  maxlength="7" size="5" class="form"> 
                                      </td>
                                      <td colspan="3" bgcolor="#FFFFFF" >&nbsp;</td>
                                    </tr>
                                    <tr> 
                                      <td bgcolor="#f5f5f5">���棺</td>
                                      <td colspan="2" bgcolor="#FFFFFF"> <input type="text" name="fax_gov_code"  maxlength="5" size="3" class="form" value="86">
                                        - 
                                        <input type="text" name="fax_area_code"  maxlength="7" size="5" class="form">
                                        - 
                                        <input type="text" name="fax"  maxlength="15" size="10" class="form">
                                         
                                        <input type="hidden" name="fax_ext_code"  maxlength="50" size="5" class="form"> 
                                      </td>
                                      <td colspan="3" bgcolor="#FFFFFF" ><font color="#CC0000">*</font></td>
                                    </tr>
                                    <tr> 
                                      <td bgcolor="#f5f5f5">OICQ/ICQ/MSN��</td>
                                      <td colspan="2" bgcolor="#FFFFFF"><input type="text" name="QQ" size="20" maxlength="40" class="form"></td>
                                      <td colspan="3" bgcolor="#FFFFFF" >&nbsp;</td>
                                    </tr>
                                    <tr> 
                                      <td bgcolor="#f5f5f5">�����ʼ���</td>
                                      <td colspan="2" bgcolor="#FFFFFF"><input type="text" name="email" size="20" maxlength="40" class="form"></td>
                                      <td colspan="3" bgcolor="#FFFFFF" ><font color="#CC0000">*</font></td>
                                    </tr>
                                    <tr> 
                                      <td bgcolor="#f5f5f5">����ע��</td>
                                      <td colspan="2" bgcolor="#FFFFFF"><textarea name="memo" cols="40" rows="5"></textarea></td>
                                      <td colspan="3" bgcolor="#FFFFFF" >&nbsp;</td>
                                    </tr>
                                    <tr> 
                                      <td colspan="6" bgcolor="#FFFFFF">&nbsp;</td>
                                    </tr>
                                    <tr bgcolor="#f5f5f5"> 
                                      <td height="35" colspan="6" align="center"> 
                                        <input class="botton" type="submit" name="ok" value="�ύ����"> 
                                        <input type="hidden" name="adduser" value="true"> 
                                        <input type="hidden" name="usertype" value="buss"></td>
                                    </tr>
                                  </form>
                                </table> </td>
                            </tr>
                            <tr> 
                              <td width="70%">&nbsp; </td>
                            </tr>
                            <tr> 
                              <td>&nbsp;</td>
                            </tr>
                          </table>
                          <br> <br> </TD>
                      </TR>
                    </TBODY>
                  </TABLE></td>
              </tr>
            </table></td>
          <td width="7" background="images/right-7x5.gif">&nbsp;</td>
        </tr>
      </table></td>
    <td width="10" background="images/dnyes-bg-left-and-right.gif"><img src="images/dnyes-bg-left-and-right.gif" width="10" height="1"></td>
  </tr>
</table>
<!--#include file="copyright.asp"-->
</body>
</html>
