<!--#include file="searchclass.asp"-->
<SCRIPT language=JavaScript>

function doValidate()
{
	if(form1.domain.value.length == 0) {
		alert("请填写一个域名！");
		form1.domain.focus();
		return false;
	}
	if(form1.domain.value.length > 26) {
		alert("域名超长！请不要超过26个，否则不能注册！");
		form1.domain.focus();
		return false;
	}
	if(! isDomain(form1.domain.value)) {
		alert("域名不合法！\n域名不能以中横杠'－'开头或结尾！\n除英文26个字母和10个阿拉伯数字以及中横杠'－'可以用作域名，其它的，请不要作域名！");
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
		alert("请填写一个域名！");
		form3.domain.focus();
		return false;
	}
	if(! ccform1(form2.domain.value)) {
		alert("首尾不能有非法字符如：-  、+、@、&等 ！");
		form2.domain.focus();
		return false;
	}
	if(! ccform2(form2.domain.value)) {
		alert("域名不能有非法字符如： #、+、@、&、=、?等！");
		form2.domain.focus();
		return false;
	}
	if(! ccform3(form2.domain.value)) {
		alert("不能是纯英文或数字域名和横杠“-”，！");
		form2.domain.focus();
		return false;
	}
		return true;
}

function checkform3(){
	if(form3.domain.value.length == 0) {
		alert("请填写一个域名！");
		form3.domain.focus();
		return false;
	}
	if(! ccform1(form3.domain.value)) {
		alert("首尾不能有非法字符如：-  、+、@、&等 ！");
		form3.domain.focus();
		return false;
	}
	if(! ccform2(form3.domain.value)) {
		alert("域名不能有非法字符如： #、+、@、&、=、?等！");
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
  domaintypestr=request("domaintypestr")   
	 
%>
<%
if trim(request("alldomain"))<>"" then 
eachdomain=split(request("alldomain"),",")
after=eachdomain(0)
domain=eachdomain(1)
domaintypestr=eachdomain(2)
end if
%>    

<% 
if domain<>"" then
	set list=new domainclass
       list.get_domain=domain&trim(after)
	   list.get_after=trim(after)
       result=list.exsit
	   exploration1=list.gettakenhtml
     set list=nothing
end if
%>

<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr><td colspan="5" height="5"></td></tr>
	
	<tr align="center"> 
      <td colspan="5"> 
<% 

if result=0 then
resultstr1=""
resultstr2="还没有被注册，点此开始注册"
else
resultstr1="已被注册"
resultstr2=""
end if

if domain<>"" then
logonstr1=resultstr1
logonstr2=resultstr2
else
logonstr1="请正确地输入要查询的域名"
logonstr12=""
end if
%>
	  </td>
	</tr>
	<tr> 
      
    <td height="25" colspan="5" align="center"><%=domaintypestr%><%=domain%><%=after%>&nbsp;<%=logonstr1%><a href="domainregtable.asp?domain=<%=domain%>&after=<%=after%>&domaintypestr=<%=domaintypestr%>" target="_blank"><font color="#FF0000"><%=logonstr2%></font></a></td>
	</tr>
	<tr> 
      
    <td height="10" colspan="5" align="center">&nbsp;</td>
	</tr>
  <form method="post" action="domain.asp"  id=form1 name=form1 onSubmit="return doValidate();">
    <tr> 
      <td  height="30" colspan="5" align="center"> <font color="#000000">英文域名查询： 
        <input type="text" name="domain" id=domain maxlength="25" size="20" class="input">
        <input type="submit" class=botton name=Submit2  id=Submit2  value=国际域名查询 align="middle" border="0" width="80" height="25">
        </font></td>
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
          <input type="radio" name="after" value=".cn">
          <font color="#FF0000">.cn</font></td>
    </tr>  
	<tr>
	<td colspan="5" align="center">
          <input type="hidden" name="domaintypestr" value="国际域名：">
      </td>
	</tr>
	</form>
	
	<tr><td colspan="5" height="25"></td></tr>
	
    <form method="post" action="domain.asp"  id=form2 name=form2 onSubmit="return checkform2();">
    <tr> 
      <td  height="30" colspan="5" align="center"> <font color="#000000">中文域名查询： 
        <input type="text" name="domain" id=domain maxlength="30" size="20" class="input">
        <input name=Submit3 type="submit" class=botton  id=Submit3 value="中文域名查询"  align="middle" width="80" height="25" border="0">
        </font></td>
    </tr>
	<tr>
      <td width="100" height="25"><p align="left"> 
          <input type="radio" name="after" value=".中国" checked>
          .中国 </td>
      <td width="100" height="25"><p align="left"> 
          <input type="radio" name="after" value=".网络">
          .网络</td>
      <td width="100" height="25"><p align="left"> 
          <input type="radio" name="after" value=".公司">
          .公司</td>
	</tr>
	<tr>
	<td colspan="5" align="center">
          <input type="hidden" name="domaintypestr" value="中文域名：">
      </td>
	</tr>
    </form>
	
	<tr><td colspan="5" height="25"></td></tr>
	
    <form method="post" action="domain.asp"  id=form3 name=form3 onSubmit="return checkform3();">
    <tr> 
    <td  height="30" colspan="5" align="center"> 
		  <div align="center"><font color="#000000">通用网址查询： 
          <input type="text" name="domain" id=domain maxlength="30" size="20" class="input">
          <input name=Submit type="submit" class=botton  id=Submit4 value="通用网址查询"  align="middle" width="80" height="25" border="0">
          </font></div>
	</td>
    </tr>

	<tr>
	<td colspan="5" align="center">
          <input type="hidden" name="domaintypestr" value="通用网址：">
      </td>
	</tr>
    </form>
	
</table>
