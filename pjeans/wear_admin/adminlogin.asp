<script language="JavaScript">

function checkform()
{
	if (form_administrator.uid.value.length == 0) {
		alert("Please enter a valid Userid. ");
		form_administrator.uid.focus();
		return false;
		}
	if(! isLetter(form_administrator.uid.value)) {
		alert('The "Userid" field makes up by "_" and "0~9" adn "a~z" and "A~Z" and blank space');
		form_administrator.uid.focus();
		return false;
		}
	if (form_administrator.pwd.value.length == 0) {
		alert(" Enter the name and type of competition. ");
		form_administrator.pwd.focus();
		return false;
		}
	if(! isLetter(form_administrator.pwd.value)) {
		alert('The "Password" field makes up by "_" and "0~9" adn "a~z" and "A~Z" and blank space');
		form_administrator.pwd.focus();
		return false;
		}
	return true;
}

function isLetter(name)
{
var letter=" _0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
	if(name.length == 0)
		return false;
	for(i = 0; i < name.length; i++) {
		if(letter.indexOf(name.charAt(i)) == -1)
			return false;
	}
	return true;
}

function isNumber(name)
{
	for(i = 0; i < name.length; i++) {
		if(name.charAt(i) < "0" || name.charAt(i) > "9")
			return false;
	}
	return true;
}

function isEnglish(name)
{
	for(i = 0; i < name.length; i++) {
		if(name.charCodeAt(i) > 128)
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

</script>
<%
if session("user_uid")<>"" then
%>
<%
if session("user_adminlevel")=1 then
%>
<table width="100%" height="100%" border="0" cellpadding="2" cellspacing="2">
        <tr> 
          <td height="410" align="center">Welcome, <font color="#FF0000">administrator</font> 
            !You have logined</td>
        </tr>
        <tr>
          
    <td height="10" align="right" valign="bottom">&nbsp;</td>
        </tr>
</table>
<%
end if
%>
<%
else
%>
<table width="100%" height="100%" border="0" cellpadding="2" cellspacing="2">
        <tr> 
		<td width="1">&nbsp;</td>
          <td align="center">
               <form action="Login.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
        <table width="168" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td height="24" valign="top"><font color="#1F4A71"><strong>User Login</strong></font></td>
          </tr>
          <tr> 
            <td>User ID : </td>
          </tr>
          <tr> 
            <td><input maxlength=20 name="uid" id="uid" size=30></td>
          </tr>
          <tr> 
            <td>Password : </td>
          </tr>
          <tr> 
            <td><input maxlength=20 name="pwd" id="pwd" size=30></td>
          </tr>
          <tr>
            <td><input name="option" type="submit" id="option" value=" Login " class="button"></td>
          </tr>
        </table>
              	</form>
				</td>
        <td width="1">&nbsp;</td>
  </tr>
</table>
<%
end if
%>
