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
		alert(" Please enter a valid PassWord. ");
		form_administrator.pwd.focus();
		return false;
		}
	if(! isLetter(form_administrator.pwd.value)) {
		alert('The "Password" field makes up by "_" and "0~9" adn "a~z" and "A~Z" and blank space');
		form_administrator.pwd.focus();
		return false;
		}
	if (form_administrator.pwd.value != form_administrator.pwdc.value) {
		alert(" Two PassWords must be equal . ");
		form_administrator.pwd.value = '';
		form_administrator.pwdc.value = '';
		form_administrator.pwd.focus();
		return false;
		}
	if(form_administrator.email.value.length > 0){
	if(! isMail(form_administrator.email.value)) {
		alert('Please enter a valid EMAIL address');
		form_administrator.email.focus();
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
              <%
sql="select * from field" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
do while not rs.eof
%>
              <tr> 
                <td width="5" valign="top"> <img src="images/sun_l_sp.gif" width="5" height="10" border="0"> 
                </td>
                <td><a href="Calendar.asp?field=<%=rs("fid")%>" title="<%=rs("name")%>" class="b"><%=rs("name")%></a></td>
              </tr>
              <%
rs.movenext
loop
rs.close
set rs=nothing
%>
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
          <td height="27">&nbsp;</td>
        </tr>
      </table>
      <table width="90%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#6184A3">
        <FORM action="Shopping_Reg_Save.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
          <tr bgcolor="#FFFFFF"> 
            <td height="30" colspan="2" align="center"><B>SIGN UP</B></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td height="30" colspan="2"># To create a new account, please enter 
              all fields below.<br>
              Already have an account? Please <a href="Index.asp"><font color="#FF0000">Login</font></a>.<br>
              <br>
              <font color="#FF0000">*</font><b>Required fields</b>.</td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="31%">User ID:</td>
            <td width="69%"><input name="uid" type="text" id="uid" size="20" maxlength="20" class="form"> 
              <font color="#FF0000"> * </font></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td>Password:</td>
            <td><input name="pwd" type="password" id="pwd" size="20" maxlength="20" class="form"> 
              <font color="#FF0000">* </font></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td>Confirm:</td>
            <td><input name="pwdc" type="password" id="pwdc" size="20" maxlength="20" class="form"> 
              <font color="#FF0000">* </font></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td>Full Name:</td>
            <td><input name="uname" type="text" id="uname" size="30" maxlength="30" class="form">
              <font color="#FF0000">*</font> </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td>Telphone:</td>
            <td><input name="tel" type="text" id="tel" size="30" maxlength="20" class="form">
              <font color="#FF0000">*</font></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td>E-mail Address:</td>
            <td><input name="email" type="text" id="email" size="30" maxlength="30" class="form">
              <br>
              (e.g., jane_doe@domain.com)<br> </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td>ZIP:</td>
            <td><input name="zip" type="text" id="zip" size="10" maxlength="10" class="form"></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td>Address:</td>
            <td><input name="contact" type="text" id="contact" size="40" maxlength="100" class="form"></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td>&nbsp;</td>
            <td> <input type="hidden" name="option" id="option" value="add"> <input type="submit" name="SING UP" value="SING UP" class="button"> 
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td colspan="2">&nbsp;</td>
          </tr>
        </form>
      </table>
	  
	  </td>
    <td width="20">&nbsp;</td>
    <td width="181" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="19">&nbsp;</td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="5" cellspacing="0">
        <tr> 
          <td><font color="#1F4A71"><strong>Why create an account?</strong></font><br>
            To check your order status, manage your Wish List and personalize 
            your shopping experience.</td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td><font color="#1F4A71"><strong>Questions?</strong></font><br>
            Call us please , SUNSPORTS (65)63448344 , 24 hours a day, 7 days a 
            week.</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
<!--#include file="Copyright.asp"-->