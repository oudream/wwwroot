<!--#include file="mymefaq5411jkjkh/favorible/showme.asp" -->
<%on error resume next%>
<script language="JavaScript">

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
		alert("�������û���.");
		document.ADDUser.uid.focus();
		return false;
	}
    if(letter.indexOf(ADDUser.uid.value.charAt(0)) == -1)
  {
  	  	alert("��Ա�ʺű���������ĸ���»�����ɣ�" )
    	ADDUser.uid.focus()
    	return false
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
		alert("��������ϵ�绰�����ţ��Ա����ǿ���Ϊ�����÷���.");
		document.ADDUser.tel1_area_code.focus();
		return false;
	}if (document.ADDUser.tel1.value.length == 0) {
		alert("��������ϵ�绰���Ա����ǿ���Ϊ�����÷���.");
		document.ADDUser.tel1.focus();
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
<title><%=Application("y_it_itname")%>-�û�ע��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="dnyes/inc/southdns.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>
			  <br>
              ע��Ϊ���ǵĻ�Ա�û������������߶������ǵĲ�Ʒ�ͷ��񲢿��Զ������Ķ��������ߵĹ���<br>���ܹ��õ����ǲ�Ʒ���±��ۣ����ǽ���Ѽ۸񣬲�Ʒ���ϵ�������Ϣ�����ڵķ��͵���������. <br>
              <br>
              ���ԣ�<font color="#FF0033">����������д��������</font>. ����<b><font color="#FF3333">*</font></b>Ϊ��ѡ��Ŀ��
	</td>
  </tr>
</table>
<table width="600" border="1" align="center">
  <form name="ADDUser" method="POST" action="userregfix.asp" onSubmit="return CheckForm();">
    <tr> 
      <td width="112">��Ա�ʺţ�</td>
      <td colspan="2"><input type="text" name="uid" maxlength="20" class="form" size="20"></td>
      <td width="108" >fsad&nbsp;</td>
      <td width="36" ><font color="#CC0000">*</font></td>
      <td width="36" >fsad&nbsp;</td>
    </tr>
    <tr> 
      <td>��Ա���룺</td>
      <td colspan="2"><input type="password" name="pwd" maxlength="20" class="form" size="20"></td>
      <td width="108" >fsad&nbsp;</td>
      <td width="36" ><font color="#CC0000">*</font></td>
      <td width="36" >fsad&nbsp;</td>
    </tr>
    <tr> 
      <td>�ظ����룺</td>
      <td colspan="2"><input type="password" name="PW_Again" maxlength="20" class="form" size="20"></td>
      <td width="108" >fsad&nbsp;</td>
      <td width="36" ><font color="#CC0000">*</font></td>
      <td width="36" >fsad&nbsp;</td>
    </tr>
    <tr> 
      <td>��ҵ���ƣ����ģ���</td>
      <td colspan="2"><input name="com_cn" type="text"  maxlength="50" ></td>
      <td width="108" >fsad&nbsp;</td>
      <td width="36" ><font color="#CC0000">*</font></td>
      <td width="36" >fsad&nbsp;</td>
    </tr>
    <tr> 
      <td>��ҵ���ƣ�Ӣ�ģ���</td>
      <td colspan="2"><input name="com_en" type="text"  maxlength="50"></td>
      <td width="108" >fsad&nbsp;</td>
      <td width="36" >&nbsp;</td>
      <td width="36" >fsad&nbsp;</td>
    </tr>

    <tr> 
      <td height="36">��ϵ�ˣ����ģ���</td>
      <td width="188"><input name="contact_cn" type="text"  size="17" maxlength="20"></td>
      <td width="156" ></td>
      <td width="108" >fsad&nbsp;</td>
      <td width="36" ><font color="#CC0000">*</font></td>
      <td width="36" >fsad&nbsp;</td>
    </tr>
    <tr> 
      <td height="20">��ϵ�ˣ�Ӣ�ģ���</td>
      <td width="188"><input name="contact_en" type="text"  size="17" maxlength="20"></td>
      <td width="108" >fsad&nbsp;</td>
      <td width="36" >&nbsp;</td>
      <td width="36" >fsad&nbsp;</td>
    </tr>
    <tr> 
      <td>��ϵ���Ա�</td>
      <td colspan="2"><input name="u_or_i" type="radio" value="1"  checked>
        �� 
        <input name="u_or_i" type="radio" value="0" >
        Ů</td>
      <td width="108" >fsad&nbsp;</td>
      <td width="36" ><font color="#CC0000">*</font></td>
      <td width="36" >fsad&nbsp;</td>
    </tr>
    <tr> 
      <td>���һ������</td>
      <td colspan="2"><select name="TECHCC"  onChange="getTECHSP()">
          <option value="CHINA" selected>�й� CHINA</option>
          <option value="AFGHANISTAN">AFGHANISTAN</option>
          <option value="ALBANIA">ALBANIA</option>
          <option value="ALGERIA">ALGERIA</option>
          <option value="AMERICAN SAMOA">AMERICAN SAMOA</option>
          <option value="ANDORRA">ANDORRA</option>
          <option value="ANGOLA">ANGOLA</option>
          <option value="ANGUILLA">ANGUILLA</option>
          <option value="ANTARCTICA">ANTARCTICA</option>
          <option value="ANTIGUA AND BARBUDA">ANTIGUA AND BARBUDA</option>
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
          <option value="BOSNIA AND HERZEGOVINA">BOSNIA AND HERZEGOVINA</option>
          <option value="BOTSWANA">BOTSWANA</option>
          <option value="BOUVET ISLAND">BOUVET ISLAND</option>
          <option value="BRAZIL">BRAZIL</option>
          <option value="BRITISH INDIAN OCEAN TERRITORY">BRITISH INDIAN OCEAN 
          TERRITORY</option>
          <option value="BRUNEI DARUSSALAM">BRUNEI DARUSSALAM</option>
          <option value="BULGARIA">BULGARIA</option>
          <option value="BURKINA FASO">BURKINA FASO</option>
          <option value="BURUNDI">BURUNDI</option>
          <option value="CAMBODIA">CAMBODIA</option>
          <option value="CAMEROON">CAMEROON</option>
          <option value="CANADA">CANADA</option>
          <option value="CAPE VERDE">CAPE VERDE</option>
          <option value="CAYMAN ISLANDS">CAYMAN ISLANDS</option>
          <option value="CENTRAL AFRICAN REPUBLIC">CENTRAL AFRICAN REPUBLIC</option>
          <option value="CHAD">CHAD</option>
          <option value="CHILE">CHILE</option>
          <option value="CHRISTMAS ISLAND">CHRISTMAS ISLAND</option>
          <option value="COCOS (KEELING) ISLANDS">COCOS (KEELING) ISLANDS</option>
          <option value="COLOMBIA">COLOMBIA</option>
          <option value="COMOROS">COMOROS</option>
          <option value="CONGO">CONGO</option>
          <option value="CONGO, THE DEMOCRATIC REPUBLIC OF THE">CONGO, THE DEMOCRATIC 
          REPUBLIC OF THE</option>
          <option value="COOK ISLANDS">COOK ISLANDS</option>
          <option value="COSTA RICA">COSTA RICA</option>
          <option value="C?TE D' IVOIR">C?TE D' IVOIRE</option>
          <option value="CROATIA">CROATIA</option>
          <option value="CUBA">CUBA</option>
          <option value="CYPRUS">CYPRUS</option>
          <option value="CZECH REPUBLIC">CZECH REPUBLIC</option>
          <option value="DENMARK">DENMARK</option>
          <option value="DJIBOUTI">DJIBOUTI</option>
          <option value="DOMINICA">DOMINICA</option>
          <option value="DOMINICAN REPUBLIC">DOMINICAN REPUBLIC</option>
          <option value="EAST TIMOR">EAST TIMOR</option>
          <option value="ECUADOR">ECUADOR</option>
          <option value="EGYPT">EGYPT</option>
          <option value="EL SALVADOR">EL SALVADOR</option>
          <option value="EQUATORIAL GUINEA">EQUATORIAL GUINEA</option>
          <option value="ERITREA">ERITREA</option>
          <option value="ESTONIA">ESTONIA</option>
          <option value="ETHIOPIA">ETHIOPIA</option>
          <option value="FALKLAND ISLANDS (MALVINAS)">FALKLAND ISLANDS (MALVINAS)</option>
          <option value="FAROE ISLANDS">FAROE ISLANDS</option>
          <option value="FIJI">FIJI</option>
          <option value="FINLAND">FINLAND</option>
          <option value="FRANCE">FRANCE</option>
          <option value="FRENCH GUIANA">FRENCH GUIANA</option>
          <option value="FRENCH POLYNESIA">FRENCH POLYNESIA</option>
          <option value="FRENCH SOUTHERN TERRITORIES">FRENCH SOUTHERN TERRITORIES</option>
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
          <option value="HEARD ISLAND AND MCDONALD ISLANDS">HEARD ISLAND AND MCDONALD 
          ISLANDS</option>
          <option value="HOLY SEE (VATICAN CITY STATE)">HOLY SEE (VATICAN CITY 
          STATE)</option>
          <option value="HONDURAS">HONDURAS</option>
          <option value="HONG KONG">HONG KONG</option>
          <option value="HUNGARY">HUNGARY</option>
          <option value="ICELAND">ICELAND</option>
          <option value="INDIA">INDIA</option>
          <option value="INDONESIA">INDONESIA</option>
          <option value="IRAN, ISLAMIC REPUBLIC OF">IRAN, ISLAMIC REPUBLIC OF</option>
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
          <option value="KOREA, DEMOCRATIC PEOPLE'S REPUBLIC OF">KOREA, DEMOCRATIC 
          PEOPLE'S REPUBLIC OF</option>
          <option value="KOREA, REPUBLIC OF">KOREA, REPUBLIC OF</option>
          <option value="KUWAIT">KUWAIT</option>
          <option value="KYRGYZSTAN">KYRGYZSTAN</option>
          <option value="LAO PEOPLE'S DEMOCRATIC REPUBLIC<">LAO PEOPLE'S DEMOCRATIC 
          REPUBLIC</option>
          <option value="LATVIA">LATVIA</option>
          <option value="LEBANON">LEBANON</option>
          <option value="LESOTHO">LESOTHO</option>
          <option value="LIBERIA">LIBERIA</option>
          <option value="LIBYAN ARAB JAMAHIRIYA">LIBYAN ARAB JAMAHIRIYA</option>
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
          <option value="MARSHALL ISLANDS">MARSHALL ISLANDS</option>
          <option value="MARTINIQUE">MARTINIQUE</option>
          <option value="MAURITANIA">MAURITANIA</option>
          <option value="MAURITIUS">MAURITIUS</option>
          <option value="MAYOTTE">MAYOTTE</option>
          <option value="MEXICO">MEXICO</option>
          <option value="MICRONESIA, FEDERATED STATES OF">MICRONESIA, FEDERATED 
          STATES OF</option>
          <option value="MOLDOVA, REPUBLIC OF">MOLDOVA, REPUBLIC OF</option>
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
          <option value="NETHERLANDS ANTILLES">NETHERLANDS ANTILLES</option>
          <option value="NEW CALEDONIA">NEW CALEDONIA</option>
          <option value="NEW ZEALAND">NEW ZEALAND</option>
          <option value="NICARAGUA">NICARAGUA</option>
          <option value="NIGER">NIGER</option>
          <option value="NIGERIA">NIGERIA</option>
          <option value="NIUE">NIUE</option>
          <option value="NORFOLK ISLAND">NORFOLK ISLAND</option>
          <option value="NORTHERN MARIANA ISLANDS">NORTHERN MARIANA ISLANDS</option>
          <option value="NORWAY">NORWAY</option>
          <option value="OMAN">OMAN</option>
          <option value="PAKISTAN">PAKISTAN</option>
          <option value="PALAU">PALAU</option>
          <option value="PALESTINIAN TERRITORY, OCCUPIED">PALESTINIAN TERRITORY, 
          OCCUPIED</option>
          <option value="PANAMA">PANAMA</option>
          <option value="PAPUA NEW GUINEA">PAPUA NEW GUINEA</option>
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
          <option value="RUSSIAN FEDERATION">RUSSIAN FEDERATION</option>
          <option value="RWANDA">RWANDA</option>
          <option value="SAINT HELENA">SAINT HELENA</option>
          <option value="SAINT KITTS AND NEVIS">SAINT KITTS AND NEVIS</option>
          <option value="SAINT LUCIA">SAINT LUCIA</option>
          <option value="SAINT PIERRE AND MIQUELON">SAINT PIERRE AND MIQUELON</option>
          <option value="SAINT VINCENT AND THE GRENADINES">SAINT VINCENT AND THE 
          GRENADINES</option>
          <option value="SAMOA">SAMOA</option>
          <option value="SAN MARINO">SAN MARINO</option>
          <option value="SAO TOME AND PRINCIPE">SAO TOME AND PRINCIPE</option>
          <option value="SAUDI ARABIA">SAUDI ARABIA</option>
          <option value="SENEGAL">SENEGAL</option>
          <option value="SEYCHELLES">SEYCHELLES</option>
          <option value="SIERRA LEONE">SIERRA LEONE</option>
          <option value="SINGAPORE">SINGAPORE</option>
          <option value="SLOVAKIA">SLOVAKIA</option>
          <option value="SLOVENIA">SLOVENIA</option>
          <option value="SOLOMON ISLANDS">SOLOMON ISLANDS</option>
          <option value="SOMALIA">SOMALIA</option>
          <option value="SOUTH AFRICA">SOUTH AFRICA</option>
          <option value="SOUTH GEORGIA AND THE SOUTH SANDWICH ISLANDS">SOUTH GEORGIA 
          AND THE SOUTH SANDWICH ISLANDS</option>
          <option value="SPAIN">SPAIN</option>
          <option value="SRI LANKA">SRI LANKA</option>
          <option value="SUDAN">SUDAN</option>
          <option value="SURINAME">SURINAME</option>
          <option value="SVALBARD AND JAN MAYEN">SVALBARD AND JAN MAYEN</option>
          <option value="SWAZILAND">SWAZILAND</option>
          <option value="SWEDEN">SWEDEN</option>
          <option value="SWITZERLAND">SWITZERLAND</option>
          <option value="SYRIAN ARAB REPUBLIC">SYRIAN ARAB REPUBLIC</option>
          <option value="CHINAESE TAIBEI">CHINAESE TAIBEI</option>
          <option value="TAJIKISTAN">TAJIKISTAN</option>
          <option value="TANZANIA, UNITED REPUBLIC OF">TANZANIA, UNITED REPUBLIC 
          OF</option>
          <option value="THAILAND">THAILAND</option>
          <option value="TOGO">TOGO</option>
          <option value="TOKELAU">TOKELAU</option>
          <option value="TONGA">TONGA</option>
          <option value="TRINIDAD AND TOBAGO">TRINIDAD AND TOBAGO</option>
          <option value="TUNISIA">TUNISIA</option>
          <option value="TURKEY">TURKEY</option>
          <option value="TURKMENISTAN">TURKMENISTAN</option>
          <option value="TURKS AND CAICOS ISLANDS">TURKS AND CAICOS ISLANDS</option>
          <option value="TUVALU">TUVALU</option>
          <option value="UGANDA">UGANDA</option>
          <option value="UKRAINE">UKRAINE</option>
          <option value="UNITED ARAB EMIRATES">UNITED ARAB EMIRATES</option>
          <option value="UNITED KINGDOM">UNITED KINGDOM</option>
          <option value="UNITED STATES">UNITED STATES</option>
          <option value="UNITED STATES MINOR OUTLYING ISLANDS">UNITED STATES MINOR 
          OUTLYING ISLANDS</option>
          <option value="URUGUAY">URUGUAY</option>
          <option value="UZBEKISTAN">UZBEKISTAN</option>
          <option value="VANUATU">VANUATU</option>
          <option value="VENEZUELA">VENEZUELA</option>
          <option value="VIET NAM">VIET NAM</option>
          <option value="VIRGIN ISLANDS, BRITISH">VIRGIN ISLANDS, BRITISH</option>
          <option value="VIRGIN ISLANDS, U.S.">VIRGIN ISLANDS, U.S.</option>
          <option value="WALLIS AND FUTUNA">WALLIS AND FUTUNA</option>
          <option value="WESTERN SAHARA">WESTERN SAHARA</option>
          <option value="YEMEN">YEMEN</option>
          <option value="YUGOSLAVIA">YUGOSLAVIA</option>
          <option value="ZAMBIA">ZAMBIA</option>
          <option value="ZIMBABWE">ZIMBABWE</option>
        </select></td>
      <td width="108" >fsad&nbsp;</td>
      <td width="36" ><font color="#CC0000">*</font></td>
      <td width="36" >fsad&nbsp;</td>
    </tr>
    <tr> 
      <td>ʡ��(����)��</td>
      <td colspan="2" id="td7" style="display:none;"><input name="TECHSP" type="text" value=""></td>
      <td width="108" id="td8"><select name="TECHSP_CN" >
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
        </select>
        <font color="#CC0000">&nbsp; </font> </td>
      <td width="36" ><font color="#CC0000">*</font></td>
      <td width="36" >fdsa</td>
    </tr>
    <tr> 
      <td>ʡ��(Ӣ��)��</td>
      <td colspan="2" id="td9" style="display:none;"> <input name="TECHEN" type="text" value="">
        <font color="#CC0000">&nbsp; </font> </td>
      <td id="td10" ><font color="#000000"><span style="font-size: 9pt"> 
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
      <td>&nbsp;</td>
      <td width="36" >fsad&nbsp;</td>
    </tr>
    <tr> 
      <td>���У����ģ���</td>
      <td colspan="2"><input type="text" name="cityz" size="17" maxlength="50" class="form"></td>
      <td width="108" >fsad&nbsp;</td>
      <td width="36" ><font color="#CC0000">*</font></td>
      <td width="36" >fsad&nbsp;</td>
    </tr>
    <tr> 
      <td>���У�Ӣ�ģ��� </td>
      <td colspan="2"><input type="text" name="citye"  maxlength="50" size="17" class="form"></td>
      <td width="108" >fsad&nbsp;</td>
      <td width="36" >&nbsp;</td>
      <td width="36" >fsad&nbsp;</td>
    </tr>
    <tr> 
      <td>ͨ�ŵ�ַ�����ģ���</td>
      <td colspan="2"><input type="text" name="addr" size="50" maxlength="100" class="form"></td>
      <td width="108" >fsad&nbsp;</td>
      <td width="36" ><font color="#CC0000">*</font></td>
      <td width="36" >fsad&nbsp;</td>
    </tr>
    <tr> 
      <td>ͨ�ŵ�ַ��Ӣ�ģ���</td>
      <td colspan="2"><input type="text" name="addr_en"  maxlength="100" size="50" class="form"></td>
      <td width="108" >fsad&nbsp;</td>
      <td width="36" >&nbsp;</td>
      <td width="36" >fsad&nbsp;</td>
    </tr>
    <tr> 
      <td>�������룺</td>
      <td colspan="2"><input type="text" name="postnumber" size="17" maxlength="50" class="form"></td>
      <td width="108" >fsad&nbsp;</td>
      <td width="36" ><font color="#CC0000">*</font></td>
      <td width="36" >fsad&nbsp;</td>
    </tr>
    <tr> 
      <td>��ϵ�绰1�� </td>
      <td colspan="2"> <input type="text" name="tel1_gov_code"  maxlength="5" size="3" class="form" value="86">
        - 
        <input type="text" name="tel1_area_code"  maxlength="50" size="5" class="form">
        - 
        <input type="text" name="tel1"  maxlength="50" size="10" class="form">
        - 
        <input type="text" name="tel1_ext_code"  maxlength="50" size="5" class="form"> 
      </td>
      <td width="108" >fsad&nbsp;</td>
      <td width="36" ><font color="#CC0000">*</font></td>
      <td width="36" >fsad&nbsp;</td>
    </tr>
    <tr> 
      <td>��ϵ�绰2�� </td>
      <td colspan="2"> <input type="text" name="tel2_gov_code"  maxlength="5" size="3" class="form" value="86">
        - 
        <input type="text" name="tel2_area_code"  maxlength="50" size="5" class="form">
        - 
        <input type="text" name="tel2"  maxlength="50" size="10" class="form">
        - 
        <input type="text" name="tel2_ext_code"  maxlength="50" size="5" class="form"> 
      </td>
      <td width="108" >fsad&nbsp;</td>
      <td width="36" >&nbsp;</td>
      <td width="36" >fsad&nbsp;</td>
    </tr>
    <tr> 
      <td>���棺</td>
      <td colspan="2"> <input type="text" name="fax_gov_code"  maxlength="5" size="3" class="form" value="86">
        - 
        <input type="text" name="fax_area_code"  maxlength="50" size="5" class="form">
        - 
        <input type="text" name="fax"  maxlength="50" size="10" class="form">
        - 
        <input type="text" name="fax_ext_code"  maxlength="50" size="5" class="form">
      </td>
      <td width="108" >fsad&nbsp;</td>
      <td width="36" >&nbsp;</td>
      <td width="36" >fsad&nbsp;</td>
    </tr>
    <tr> 
      <td>OICQ/ICQ/MSN��</td>
      <td colspan="2"><input type="text" name="QQ" size="17" maxlength="15" class="form"></td>
      <td width="108" >fsad&nbsp;</td>
      <td width="36" >&nbsp;</td>
      <td width="36" >fsad&nbsp;</td>
    </tr>
    <tr> 
      <td>�����ʼ���</td>
      <td colspan="2"><input type="text" name="email" size="17" maxlength="30" class="form"></td>
      <td width="108" >fsad&nbsp;</td>
      <td width="36" ><font color="#CC0000">*</font></td>
      <td width="36" >fsad&nbsp;</td>
    </tr>
    <tr> 
      <td>����ע��</td>
      <td colspan="2"><textarea name="memo" cols="35" rows="5"></textarea></td>
      <td width="108" >fsad&nbsp;</td>
      <td width="36" >&nbsp;</td>
      <td width="36" >fsad&nbsp;</td>
	</tr>
    <tr> 
      <td colspan="3" align="center"> <input type="submit" name="ok" value="�ύ����"> 
        <input type="hidden" name="adduser" value="true"> </td>
      <td width="108" >fsad&nbsp;</td>
      <td width="36" >&nbsp;</td>
      <td width="36" >fsad&nbsp;</td>
    </tr>
  </form>
</table>
</body>
</html>
