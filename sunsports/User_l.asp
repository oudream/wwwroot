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
<table width="174" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          
    <td height="15" colspan="3"><img src="1.gif" width="1" height="1"></td>
        </tr>
<%
if session("user_uid")<>"" then
if session("user_adminlevel")=1 then
%>
        <tr valign="top"> 
          
    <td colspan="3"><TABLE width=168 border=0 align="center" cellpadding="0" cellspacing="0">
        <TR> 
          <TD height="24"><font color="#1F4A71"><strong>Administrator Login : 
            : : : :</strong></font></TD>
        </TR>
        <TR> 
        <TR> 
          <TD> Dear : <%=session("user_uid")%></TD>
        </TR>
        <TR> 
          <TD valign="middle"><a href="admin/index.asp" target="_blank">Control 
            panel here</a></TD>
        </TR>
        <TR> 
          <TD valign="middle"><a href="Shopping_Buy_View.asp">Manage Your Orders</a></TD>
        </TR>
        <TR> 
          <TD valign="middle"><a href="Shopping_Reg_Epwd.asp">Manage Your Password</a></TD>
        </TR>
        <TR> 
          <TD valign="middle"><a href="Shopping_Reg_Edit.asp">Manage Your Info</a></TD>
        </TR>
        <TR>
          <TD valign="middle">&nbsp;</TD>
        </TR>
      </TABLE> </td>
        </tr>
<%
elseif session("user_adminlevel")=2 then
%>
        <tr valign="top"> 
          <td colspan="3">
              <TABLE width=168 border=0 align="center" cellpadding="0" cellspacing="0">
        <TR> 
          <TD height="24"> <strong><font color="#1F4A71">Manager Login : : : : 
            :</font></strong></TD>
        </TR>
        <TR> 
          <TD> Dear : <%=session("user_uid")%></TD>
        </TR>
        <TR> 
          <TD valign="middle"><a href="manager/index.asp" target="_blank">Control 
            panel here</a></TD>
        </TR>
        <TR> 
          <TD valign="middle"><a href="Shopping_Buy_View.asp">Manage Your Orders</a></TD>
        </TR>
        <TR> 
          <TD valign="middle"><a href="Shopping_Reg_Epwd.asp">Manage Your Password</a></TD>
        </TR>
        <TR> 
          <TD valign="middle"><a href="Shopping_Reg_Edit.asp">Manage Your Info</a></TD>
        </TR>
        <TR>
          <TD valign="middle">&nbsp;</TD>
        </TR>
		
      </TABLE>		  
	  </td>
        </tr>

<%
elseif session("user_adminlevel")=3 or session("user_adminlevel")=4 or session("user_adminlevel")=5 then
%>
<!-- 
   ----------------------------------------------   ----------------------------------------------------
   														logined level3																
   ----------------------------------------------   ----------------------------------------------------
-->
  <tr> 
    <td>&nbsp;</td>
	    <td valign="top"> 
            <TABLE width=168 border=0 align="center" cellpadding="0" cellspacing="0">
        <TR> 
          <TD height="24"><font color="#1F4A71"><strong>User Login : : : : :</strong></font></TD>
        </TR>
        <TR> 
          <TD>User : <%=session("user_uid")%></TD>
        </TR>
        <TR> 
          <TD valign="middle"><a href="Shopping_Buy_View.asp">Manage Your Orders</a></TD>
        </TR>
        <TR> 
          <TD valign="middle"><a href="Shopping_Reg_Epwd.asp">Manage Your Password</a></TD>
        </TR>
        <TR> 
          <TD valign="middle"><a href="Shopping_Reg_Edit.asp">Manage Your Info</a></TD>
        </TR>
        <TR>
          <TD valign="middle">&nbsp;</TD>
        </TR>
      </TABLE>
					  </td>
    <td>&nbsp;</td>
  </tr>
<%
else response.Redirect("error.asp?err=001")
end if
%>
<!-- 
   ----------------------------------------------   ----------------------------------------------------
   														loginout																
   ----------------------------------------------   ----------------------------------------------------
-->
  <tr>
    <td>&nbsp;</td>
    <td>
	
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>      <FORM action=Login.asp method="post" id="form_administrator" name="form_administrator" onSubmit="return checkform();">
 
          <td width="58">
            <INPUT type=submit value="logout" id="option" name="option" class="button">
		  </td>
          <td width="114">&nbsp;</td>          </FORM>

        </tr>
      </table> 
	</td>
    <td>&nbsp;</td>
  </tr>
<%
else
%>
<!-- 
   ----------------------------------------------   ----------------------------------------------------
   														login																
   ----------------------------------------------   ----------------------------------------------------
-->
        <tr> 
		<td width="1">&nbsp;</td>
          <td align="center">
               <form action="Login.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
				
            
        <table width="168" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td height="24" valign="top"><font color="#1F4A71"><strong>User Login 
              : : : : :</strong></font></td>
          </tr>
          <tr> 
            <td>User ID : </td>
          </tr>
          <tr> 
            <td><input maxlength=20 name="uid" id="uid" size=30 class="input" style="FONT-SIZE: 12px; WIDTH: 150px; background:white;border:1px solid #6184A3" onMouseOver="this.style.border='1px solid red'" onMouseOut="this.style.border='1px solid #6184A3'"></td>
          </tr>
          <tr> 
            <td>Password : </td>
          </tr>
          <tr> 
            <td><input maxlength=20 name="pwd" id="pwd" size=30 type=password class="input" style="FONT-SIZE: 12px; WIDTH: 150px; background:white;border:1px solid #6184A3" onMouseOver="this.style.border='1px solid red'" onMouseOut="this.style.border='1px solid #6184A3'"></td>
          </tr>
          <tr> 
            <td height="10"><img src="1.gif" width="1" height="1"></td>
          </tr>
          <tr>
            <td><input name="option" type="submit" id="option" value=" Login " class="button"> 
              <input name="url" id="url" type="hidden" value="index.asp"> <a href="Shopping_reg.asp" target="_blank">SIGN 
              UP</a></td>
          </tr>
        </table>
              	</form>
				</td>
        <td width="1">&nbsp;</td>
  </tr>
<%
end if
%>
</table>