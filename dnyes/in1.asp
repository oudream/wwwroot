<!--#include file="whoisclass.asp"-->
<SCRIPT language=JavaScript>
function doValidate()
{
	if(form1.domain.value.length == 0) {
		alert("����дһ��������");
		form1.domain.focus();
		return false;
	}
	if(form1.domain.value.length > 26) {
		alert("�����������벻Ҫ����26����������ע�ᣡ");
		form1.domain.focus();
		return false;
	}
	if(! isDomain(form1.domain.value)) {
		alert("�������Ϸ���\n�����������к��'��'��ͷ���β��\n��Ӣ��26����ĸ��10�������������Լ��к��'��'�������������������ģ��벻Ҫ��������");
		form1.domain.focus();
		return false;
	}
	return true;
}

function isDomain(name)
{
	if(name.length > 0) {
		if(name.charAt(0) == "-" || name.charAt(name.length - 1) == "-")
			return false;
	}
	for(i = 0; i < name.length; i++) {
		if((name.charAt(i) < "a" || name.charAt(i) > "z") && (name.charAt(i) < "A" || name.charAt(i) > "Z") && (name.charAt(i) < "0" || name.charAt(i) > "9") && name.charAt(i) != "-")
			return false;
	}
	return true;
}


function checkform2(){
	if(form2.domain.value.length == 0) {
		alert("����дһ��������");
		form3.domain.focus();
		return false;
	}
	if(! ccform1(form2.domain.value)) {
		alert("��β�����зǷ��ַ��磺-  ��+��@��&�� ��");
		form2.domain.focus();
		return false;
	}
	if(! ccform2(form2.domain.value)) {
		alert("���������зǷ��ַ��磺 #��+��@��&��=��?�ȣ�");
		form2.domain.focus();
		return false;
	}
	if(! ccform3(form2.domain.value)) {
		alert("�����Ǵ�Ӣ�Ļ����������ͺ�ܡ�-������");
		form2.domain.focus();
		return false;
	}
		return true;
}

function checkform3(){
	if(form3.domain.value.length == 0) {
		alert("����дһ��������");
		form3.domain.focus();
		return false;
	}
	if(! ccform1(form3.domain.value)) {
		alert("��β�����зǷ��ַ��磺-  ��+��@��&�� ��");
		form3.domain.focus();
		return false;
	}
	if(! ccform2(form3.domain.value)) {
		alert("���������зǷ��ַ��磺 #��+��@��&��=��?�ȣ�");
		form3.domain.focus();
		return false;
	}
		return true;
}

function ccform1(ccf)
{
	if(ccf.length > 0) {
		if(ccf.charAt(0) == "-" || ccf.charAt(ccf.length - 1) == "-"){
	return false;}
	}
	return true;
}

function ccform2(ccf)
{
	for(i = 0; i < ccf.length; i++) {
		if((ccf.charAt(i) < "-" && ccf.charAt(i) >= "!") || (ccf.charAt(i) <= "/" && ccf.charAt(i) > "-")|| (ccf.charAt(i) <= "@" && ccf.charAt(i) >= ":") || (ccf.charAt(i) <= "_" && ccf.charAt(i) >= "[") ||(ccf.charAt(i) <= "}" && ccf.charAt(i) >= "{")){
	return false;}
	}
	return true;
}

function ccform3(ccf)
{
	var letter =  "_0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
	s=0;
	for(i = 0; i < ccf.length; i++) {
	    if(letter.indexOf(ccf.charAt(i)) != -1){
		s=s+i ;}
	}
	
	sccf=0;
	for(i = 0; i < ccf.length; i++) {
		sccf=sccf+i;}
		
	if (s==sccf){return false;}
	
  	return true;
}


</SCRIPT>

<% on error resume next 
  dim after,domain,result
  after=request.form("after")
  domain=trim(request.form("domain"))

if domain<>"" then   
     set list=new whoisclass
       list.get_domain=domain&trim(after)
	   list.get_after=trim(after)
       result=list.exsit
	   exploration=list.gettakenhtml
     set list=nothing
end if	 
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <form method="post" action="showwhois.asp" onSubmit="return doValidate();" id=form1 name=form1  target="_blank">
    <tr> 
    <td  height="25" colspan="5" align="center"> 
		  <div align="center"><font color="#000000">Ӣ��������ѯ�� 
          <input type="text" name="domain" id=domain maxlength="30" size="20" class="input">
          <input class="botton" type="submit" name=Submit2  id=Submit2  align="middle" border="0" width="80" height="25">
          </font></div>
	</td>
    </tr>
    <tr> 
      <td width="100" height="25"><p align="left"> 
          <input type="radio" name="after" value=".com" checked>
          .com </td>
      <td width="100" height="25"><p align="left"> 
          <input type="radio" name="after" value=".net">
          .net </td>
      <td width="100" height="25"><p align="left"> 
          <input type="radio" name="after" value=".org">
          .org </td>
      <td width="100" height="25"><p align="left"> 
          <input type="radio" name="after" value=".biz">
          .biz </td>
      <td width="100" height="25"><p align="left"> 
          <input type="radio" name="after" value=".info">
          .info </td>
    </tr>
    <tr> 
      <td width="100" height="25"><p align="left"> 
          <input type="radio" name="after" value=".com.cn">
          .com.cn </td>
      <td width="100" height="25"><p align="left"> 
          <input type="radio" name="after" value=".net.cn">
          .net.cn </td>
      <td width="100" height="25"><p align="left"> 
          <input type="radio" name="after" value=".org.cn">
          .org.cn </td>
      <td width="100" height="25"><p align="left"> 
          <input type="radio" name="after" value=".gov.cn">
          .gov.cn</td>
      <td width="100" height="25"><p align="left"> 
          <input type="hidden" name="domaintypestr" value="����������ѯ">
          <input type="radio" name="after" value=".cn">
          .cn</td>
    </tr>  
	<tr>
	  <td colspan="5" align="center"> 
        <!-- û����ͼ��� ---------------------------------------------------------------------------?????-->
        <!-- û����ͼ��� ---------------------------------------------------------------------------?????-->
      </td>
	</tr>
	</form>
	
	<tr><td colspan="5" height="5"></td></tr>
	
    <form method="post" action="showwhois.asp"  id=form2  name=form2  onSubmit="return checkform2();" target="_blank">
    <tr> 
    <td  height="25" colspan="5" align="center"> 
		  <div align="center"><font color="#000000">����������ѯ�� 
          <input type="text" name="domain" id=domain maxlength="16" size="20" class="input">
          <input class="botton" type="submit" name=Submit3  id=Submit3  align="middle" border="0" width="80" height="25">
          </font></div>
	</td>
    </tr>
	<tr>
      <td width="100" height="25"><p align="left"> 
          <input type="radio" name="after" value=".�й�" checked>
          .�й� </td>
      <td width="100" height="25"><p align="left"> 
          <input type="radio" name="after" value=".����">
          .����</td>
      <td width="100" height="25"><p align="left"> 
          <input type="hidden" name="domaintypestr" value="����������ѯ">
          <input type="radio" name="after" value=".��˾">
          .��˾</td>
	</tr>
	<tr>
	  <td colspan="5" align="center"> 
        <!-- û����ͼ��� ---------------------------------------------------------------------------?????-->
        <!-- û����ͼ��� ---------------------------------------------------------------------------?????-->
      </td>
	</tr>
    </form>
	
	<tr><td colspan="5" height="5"></td></tr>
	
    <form method="post" action="showwhois.asp"  id=form3 name=form3 onSubmit="return checkform3();" target="_blank">
    <tr> 
    <td  height="25" colspan="5" align="center"> 
		  <div align="center"><font color="#000000">ͨ����ַ��ѯ�� 
          <input type="hidden" name="domaintypestr" value="ͨ����ַ��ѯ">
          <input type="text" name="domain" id=domain maxlength="16" size="20" class="input">
          <input class="botton" type="submit" name=Submit  id=Submit5  align="middle" border="0" width="80" height="25">
          </font></div>
	</td>
    </tr>

	<tr>
	  <td colspan="5" align="center"> 
      </td>
	</tr>
    </form>
	
	<tr><td colspan="5" height="5"></td></tr>
	
	<tr> 
      <td colspan="5"> 
	  </td>
	</tr>
</table>
