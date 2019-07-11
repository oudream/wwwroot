<script language="JavaScript">

function checkform_usermod()
{
	if (form_usermod.oldpwd.value.length == 0) {
		alert(" Please enter a valid Old PassWord. ");
		form_usermod.oldpwd.focus();
		return false;
		}
	if(! isLetter(form_usermod.oldpwd.value)) {
		alert('The "Old Password" field makes up by "_" and "0~9" adn "a~z" and "A~Z" and blank space');
		form_usermod.oldpwd.focus();
		return false;
		}
	if (form_usermod.pwd.value.length == 0) {
		alert(" Please enter a valid PassWord. ");
		form_usermod.pwd.focus();
		return false;
		}
	if(! isLetter(form_usermod.pwd.value)) {
		alert('The "Password" field makes up by "_" and "0~9" adn "a~z" and "A~Z" and blank space');
		form_usermod.pwd.focus();
		return false;
		}
	if (form_usermod.pwd.value != form_usermod.pwdc.value) {
		alert(" Two PassWords must be equal . ");
		form_usermod.pwd.value = '';
		form_usermod.pwdc.value = '';
		form_usermod.pwd.focus();
		return false;
		}
	if(form_usermod.email.value.length > 0){
	if(! isMail(form_usermod.email.value)) {
		alert('Please enter a valid EMAIL address');
		form_usermod.email.focus();
		return false;
		}}
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
<!--#include file="Top_sun.asp"-->
<table width="776" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="181" valign="top"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td> 
            <!-- ========================================================================================================		  

													cup table ============star

 ======================================================================================================== -->
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="League.asp"><b><font color="#1F4A71">LEAGUE MATCHES</font></b></a><font color="#1F4A71">+ 
                  +</font></td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="Cup.asp"><b><font color="#1F4A71">CUP MATCHES </font></b></a><font color="#1F4A71">&nbsp;+ 
                  +</font></td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="Friendlies.asp"><b><font color="#1F4A71">FRIENDLY 
                  MATCHES </font></b></a></td>
              </tr>
            </table>
            <table width="169" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td width="163"><a href="TeamProfile.asp"><b><font color="#1F4A71">TEAMPORFILE</font></b></a> 
                  <font color="#1F4A71">&nbsp;+ +</font> </td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="TeamProfile.asp"><b><font color="#1F4A71">FIELDS 
                  BOOKING</font></b></a></td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="Shopping_sort_view.asp"><b><font color="#1F4A71">SUNSPORTS 
                  PRODUCTS </font></b></a> <font color="#1F4A71">+ +</font></td>
              </tr>
            </table>
            <!-- ========================================================================================================		  

													cup table ============stop

 ======================================================================================================== -->
          </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table></td>
    <td width="20">&nbsp; </td>
    <td valign="top"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="25">&nbsp;</td>
        </tr>
      </table>
      <table width="90%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#6184A3">
        <FORM action="Shopping_Reg_Save.asp" method="post" name="form_usermod" id="form_usermod" onSubmit="return checkform_usermod();">
          <tr align="center" bgcolor="#FFFFFF"> 
            <td height="30" colspan="2"><B>Change Your Password</B></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="28%">Dear :</td>
            <td width="72%"><%=session("user_uid")%>&nbsp;</td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td>Old Password:</td>
            <td><input name="oldpwd" type="password" id="oldpwd" size="20" maxlength="20" class="form"> 
              <font color="#FF0000">* </font></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td>New Password:</td>
            <td><input name="pwd" type="password" id="pwd" size="20" maxlength="20" class="form"> 
              <font color="#FF0000">* </font></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td>Confirm:</td>
            <td><input name="pwdc" type="password" id="pwdc" size="20" maxlength="20" class="form"> 
              <font color="#FF0000">* </font></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td>&nbsp;</td>
            <td> 
			<input type="hidden" name="option" id="option" value="editpwd"> 
			<input type="submit" name="SING UP" value="Edit" class="button"> 
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td colspan="2"> # We recommend that you periodically update your 
              password to help keep your information secure. Change it by entering 
              the required information below. <br>
              <br>
              <br>
              # <font color="#FF0000">*</font>Required Fields<br>
              <br>
              <br>
            </td>
          </tr>
        </form>
      </table>
	  </td>
    <td width="20" valign="top"><table width="95%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="24">&nbsp;</td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td>&nbsp;</td>
          <td height="280" background="images/10x1_blue.gif">&nbsp;</td>
        </tr>
      </table></td>
    <td width="181" valign="top">
          <!--#include file="User_l.asp" -->
	 <br> <!--#include file="Shopping_Search.asp" -->
	 <br> 	
	</td>
  </tr>
</table>
<!--#include file="Copyright.asp"-->