<script language="JavaScript">

function checkform()
{
	if (form_administrator.uname.value.length == 0) {
		alert("Please enter a valid Name. ");
		form_administrator.uname.focus();
		return false;
		}
	if(! isLetter(form_administrator.uname.value)) {
		alert('The "Name" field makes up by "_" and "0~9" adn "a~z" and "A~Z" and blank space');
		form_administrator.uname.focus();
		return false;
		}
	if(! isMail(form_administrator.email.value)) {
		alert('Please enter a valid EMAIL address');
		form_administrator.email.focus();
		return false;
		}
	if(! isNumber(form_administrator.nric.value)) {
		alert('Use integers only for "Nric/Passort No." field');
		form_administrator.nric.focus();
		return false;
		}
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

<!--#include file="top.asp" -->
<%
adminlevel=trim(request("adminlevel"))
if adminlevel="" then response.Redirect("error.asp?err=001")
uname=trim(request("uname"))
contact=trim(request("contact"))
email=trim(request("email"))
nric=trim(request("nric"))
uid=trim(request("uid"))
pwd=trim(request("pwd"))
if contact="" then contact=" "
if email="" then email=" "
if nric="" then nric=" "
if uname<>"" then
sql="insert into user (name,contact,email,uid,pwd,nric,adminlevel,inserttime) values ('"&uname&"','"&contact&"','"&email&"','"&uid&"','"&pwd&"','"&nric&"',"&adminlevel&",'"&now()&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Create administrator is complete. ');</SCRIPT>")
end if
%>
<H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>Create Administrator</FONT></H2>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<FORM action="../administrator_c.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
  <TABLE height=307 cellSpacing=0 cellPadding=0 width=500 
                  border=0>
    <TBODY>
      <TR>
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              NAME
<!------------------------------------------------------------------------------------------------------------------------------------ -->
        <TD width=104 height=45> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3>Name</FONT></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3>
          <INPUT maxLength=100 size=30 id="uname" name="uname">
          *</FONT></TD>
      </TR>
      <TR>
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              contact
<!------------------------------------------------------------------------------------------------------------------------------------ -->
        <TD width=104 height=45> <DIV align=center><font color="#000000" size="3" face="Verdana, Arial, Helvetica, sans-serif">Contact</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3>
          <INPUT maxLength=100 size=30 id="contact" name="contact">
          *</FONT></TD>
      </TR>
      <TR>
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
        <TD width=104 height=45> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3>Email</FONT></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3>
          <INPUT maxLength=40 size=30 id="email" name="email">
          *</FONT></TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							             NRIC
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD height=43> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3>Nric/Passort No.</FONT></DIV></TD>
        <TD><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3>
          <INPUT maxLength=10 size=30 id="nric" name="nric">
          * </FONT></TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              USERID
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD height=42> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3>Userid</FONT></DIV></TD>
        <TD><INPUT maxLength=20 size=30 id="uid" name="uid">
		 <FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3>*</FONT></TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							         PASSWORD
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD height=32> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3>Password</FONT></DIV></TD>
        <TD><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3>
          <INPUT maxLength=20 size=30 id="pwd" name="pwd">
          * </FONT></TD>
      </TR>
      <TR> 
        <TD colSpan=2> <P><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2><STRONG>* Mandatory Fields</STRONG></FONT></P>
          <P><STRONG><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2>
            <INPUT type=submit value=Submit name=Submit>
            </FONT></STRONG></P></TD>
      </TR>
    </TBODY>
  </TABLE>
  <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							HIDDEN FIELD
<!------------------------------------------------------------------------------------------------------------------------------------ -->
  			<INPUT name="adminlevel" type=hidden id="adminlevel" value=<%=request("adminlevel")%>>
</FORM>
<P><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066></FONT></P>
<P><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066></FONT></P>
<%=uname%><br>
<%=contact%><br>
<%=email%><br>
<%=nric%><br>
<%=uid%><br>
<%=pwd%><br>
<%=adminlevel%>