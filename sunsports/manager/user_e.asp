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

<!--#include file="adminconn.asp" -->
<%
adminlevel=request("adminlevel")
id=request("id")
page=request("page")
if adminlevel="" or id="" or page="" then response.Redirect("error.asp?err=001")
%>
<%
if request("option")="del" then
if adminlevel=1 or adminlevel=3 then conn.execute "delete * from player where pid=" & id
if adminlevel=2 then conn.execute "delete * from player where pid=" & id
url="user_v.asp?page=" & page & "&adminlevel=" & adminlevel
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Del administrator is complete. ');</SCRIPT>")
response.redirect url
end if
%>
<%
if request("option")="edit" then
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
if adminlevel=1 or adminlevel=3 then sql="update player set name='"&uname&"',contact='"&contact&"',email='"&email&"',nric='"&nric&"',uid='"&uid&"',pwd='"&pwd&"' where pid=" & id
if adminlevel=2 then sql="update player set name='"&uname&"',contact='"&contact&"',email='"&email&"',nric='"&nric&"',uid='"&uid&"',pwd='"&pwd&"' where pid=" & id
conn.Execute(sql)
url="user_v.asp?page=" & page & "&adminlevel=" & adminlevel
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Create administrator is complete. ');</SCRIPT>")
response.redirect url
end if
end if
%>
<%
if adminlevel=1 or adminlevel=3 then sql="select * from player where pid=" & id
if adminlevel=2 then sql="select * from player where pid=" & id
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
%>
<H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>Edit Administrator</FONT></H2>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<FORM action="user_e.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
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
          <INPUT maxLength=100 size=30 id="uname" name="uname" value="<%=rs("name")%>">
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
          <INPUT maxLength=100 size=30 id="contact" name="contact" value="<%=rs("contact")%>">
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
          <INPUT maxLength=40 size=30 id="email" name="email" value="<%=rs("email")%>">
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
          <INPUT maxLength=10 size=30 id="nric" name="nric"  value="<%=rs("nric")%>">
          * </FONT></TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              USERID
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD height=42> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3>Userid</FONT></DIV></TD>
        <TD><INPUT maxLength=20 size=30 id="uid" name="uid"  value="<%=rs("uid")%>">
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
          <INPUT maxLength=20 size=30 id="pwd" name="pwd" value="<%=rs("pwd")%>">
          * </FONT></TD>
      </TR>
      <TR> 
        <TD colSpan=2> <P><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2><STRONG>* Mandatory Fields</STRONG></FONT></P>
          <P><STRONG><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2>
            <INPUT type=submit id="option" name="option" value="edit">&nbsp;&nbsp;
            <INPUT type=submit id="option" name="option" value="del" onClick="return confirm('Are you sure del it ?')">&nbsp;&nbsp;
            <INPUT type="button" id="goback" name="goback" value="go back" onClick="history.go( -1 );">
            </FONT></STRONG></P></TD>
      </TR>
    </TBODY>
  </TABLE>
  <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							HIDDEN FIELD
<!------------------------------------------------------------------------------------------------------------------------------------ -->
  			<INPUT name="adminlevel" type=hidden id="adminlevel" value=<%=adminlevel%>>
  			<INPUT name="id" type=hidden id="id" value=<%=id%>>
  			<INPUT name="page" type=hidden id="page" value=<%=page%>>
</FORM>
<%
rs.close
set rs=nothing
%>
<P><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066></FONT></P>
<P><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066></FONT></P>
