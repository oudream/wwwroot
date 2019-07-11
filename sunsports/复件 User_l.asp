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
<table width="39%" border="0" cellspacing="2" cellpadding="2">
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="45%">&nbsp;</td>
    <td width="46%">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td><table width="98%" border="0" cellspacing="0" cellpadding="0">
        <tr align="center"> 
          <td height="15" colspan="3">dsfsda</td>
        </tr>
<%
if session("user_uid")<>"" then
if session("user_adminlevel")=1 then
%>
        <tr align="center"> 
          <td height="15" colspan="3">
		<TABLE height=200 width=100% border=0>
              <TR> 
                
              <TD vAlign=middle align=left height=200><TABLE width=100 border=0>
                  <TR> 
                    <TD>&nbsp;</TD>
                  </TR>
                </TABLE>
                <TABLE width=100% border=0>
                  <TR> 
                    <TD width=131> Dear : <%=session("user_uid")%></TD>
                  </TR>
                  <TR>
                    <TD align="center" valign="middle"><a href="admin/index.asp">Here to administrator</a></TD>
                  </TR>
                </TABLE>
                <TABLE width=100 border=0>
                  <TR> 
                    <TD>&nbsp;</TD>
                  </TR>
                </TABLE> </TD>
            </TR>
        </TABLE>
		  </td>
        </tr>
<%
elseif session("user_adminlevel")=2 then
%>
        <tr align="center"> 
          <td height="15" colspan="3">
		<TABLE height=200 width=100% border=0>
              <TR> 
                
              <TD vAlign=middle align=left height=200><TABLE width=100 border=0>
                  <TR> 
                    <TD>&nbsp;</TD>
                  </TR>
                </TABLE>
                <TABLE width=100% border=0>
                  <TR> 
                    <TD width=131 align="center"> Manager</TD>
                  </TR>
                   <TR> 
                    <TD width=131> Dear : <%=session("user_uid")%></TD>
                  </TR>
                 <TR>
                    <TD align="center" valign="middle"><a href="manager/index.asp">Here to administrator</a></TD>
                  </TR>
                </TABLE>
                <TABLE width=100 border=0>
                  <TR> 
                    <TD>&nbsp;</TD>
                  </TR>
                </TABLE> </TD>
            </TR>
        </TABLE>
		  </td>
        </tr>

<%
elseif session("user_adminlevel")=3 then
%>
<!-- 
   ----------------------------------------------   ----------------------------------------------------
   														logined level3																
   ----------------------------------------------   ----------------------------------------------------
-->
  <tr> 
    <td>&nbsp;</td>
	    <td align="center"> 
		<TABLE height=200 width=100% border=0>
              <TR> 
                
              <TD vAlign=middle align=left height=200><TABLE width=100 border=0>
                  <TR> 
                    <TD>&nbsp;</TD>
                  </TR>
                </TABLE>
                <TABLE width=180 border=0>
                  <TR> 
                    <TD width=131> Dear : You have logined</TD>
                  </TR>
                  <TR>
                    <TD>You can ...</TD>
                  </TR>
                </TABLE>
                <TABLE width=100 border=0>
                  <TR> 
                    <TD>&nbsp;</TD>
                  </TR>
                </TABLE> </TD>
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
          <FORM action=Login.asp method="post" id="form_administrator" name="form_administrator" onSubmit="return checkform();">
            <INPUT type=submit value="logout" id="option" name="option">
          </FORM>
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
          <td width="209" align="center">
               <form action="Login.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
				
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td height="30">&nbsp;</td>
              </tr>
              <tr> 
                <td height="23">Userid : </td>
              </tr>
              <tr align="right"> 
                <td height="29"><input maxlength=20 name="uid" id="uid" size=20 class="input"></td>
              </tr>
              <tr> 
                <td height="26">Password : </td>
              </tr>
              <tr> 
                <td height="33" align="right"><input maxlength=20 name="pwd" id="pwd" size=20 type=password class="input"></td>
              </tr>
              <tr> 
                <td height="30" align="center"><input name="option" type="submit" id="option" value="login"> 
                  <input name="url" id="url" type="hidden" value="index.asp">
                </td>
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
<br>