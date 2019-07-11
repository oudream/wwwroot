<script language="JavaScript">

function checkform()
{
	if (form_administrator.uname.value.length == 0) {
		alert("Please enter a valid Name. ");
		form_administrator.uname.focus();
		return false;
		}
	if(! isEnglish(form_administrator.uname.value)) {
		alert('Please enter a valid Name.');
		form_administrator.uname.focus();
		return false;
		}
	if(form_administrator.email.value.length > 0){
	if(! isMail(form_administrator.email.value)) {
		alert('Please enter a valid EMAIL address');
		form_administrator.email.focus();
		return false;
		}}
	if(! isNumber(form_administrator.jersey.value)) {
		alert('Use integers only for "jersey" field');
		form_administrator.jersey.focus();
		return false;
		}
	if (form_administrator.uid.value.length > 1) {
	if(! isLetter(form_administrator.uid.value)) {
		alert('The "Userid" field makes up by "_" and "0~9" adn "a~z" and "A~Z" and blank space');
		form_administrator.uid.focus();
		return false;
		}
		}
	if (form_administrator.pwd.value.length > 1) {
	if(! isLetter(form_administrator.pwd.value)) {
		alert('The "Password" field makes up by "_" and "0~9" adn "a~z" and "A~Z" and blank space');
		form_administrator.pwd.focus();
		return false;
		}
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
if request("option")="add" then
uname=trim(request("uname"))
uname=replace(uname,"'","''")

uid=trim(request("uid"))
pwd=trim(request("pwd"))
if uid="" then 
zuid="-1"
else
zuid=uid
end if
usql="select * from player where uid='"&zuid&"'"
set urs=server.createobject("ADODB.Recordset")
urs.open usql,conn,1,1
if not ( urs.eof or urs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' The Userid have existed , please enter other Userid . ');</SCRIPT>")
adminlevel=0
urs.close
set urs=nothing
else
urs.close
set urs=nothing

tel=trim(request("tel"))
tel=replace(tel,"'","''")
zip=trim(request("zip"))
zip=replace(zip,"'","''")
contact=trim(request("contact"))
contact=replace(contact,"'","''")
email=trim(request("email"))
email=replace(email,"'","''")
nric=trim(request("nric"))
nric=replace(nric,"'","''")
jersey=trim(request("jersey"))
default_score=trim(request("default_score"))
default_yellow=trim(request("default_yellow"))
default_red=trim(request("default_red"))
if uid="" then uid=" "
if pwd="" then pwd=" "
if tel="" then tel=" "
if zip="" then zip=" "
if contact="" then contact=" "
if email="" then email=" "
if nric="" then nric=" "
if jersey="" then jersey=0
if default_score="" then default_score=0
if default_yellow="" then default_yellow=0
if default_red="" then default_red=0
sql="insert into player (name,uid,pwd,tel,zip,contact,email,nric,ptype,jersey,default_score,default_yellow,default_red,inserttime,adminlevel) values ('"&uname&"','"&uid&"','"&pwd&"','"&tel&"','"&zip&"','"&contact&"','"&email&"','"&nric&"','"&"a"&"',"&jersey&","&default_score&","&default_yellow&","&default_red&",'"&now()&"',5)"
conn.Execute(sql)
response.Redirect("team_c.asp?option=creammangerasstsucc")
end if
end if
%>
<H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>Create Asst.Manager</FONT></H2>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<FORM action="manager_asst_c.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
  <TABLE height=307 cellSpacing=0 cellPadding=0 width=500 
                  border=0>
    <TBODY>
      <TR> 
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              NAME
<!------------------------------------------------------------------------------------------------------------------------------------ -->
        <TD width=193 height=45> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3>Name</FONT></DIV></TD>
        <TD width=307><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3> 
          <INPUT maxLength=100 size=30 id="uname" name="uname">
          <font color="#FF0000">*</font></FONT></TD>
      </TR>
      <TR> 
        <TD width=104 height=45> <DIV align=center><font color="#000000" size="3" face="Verdana, Arial, Helvetica, sans-serif">Telphone</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3> 
          <INPUT maxLength=100 size=30 id="tel" name="tel">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=104 height=45> <DIV align=center><font color="#000000" size="3" face="Verdana, Arial, Helvetica, sans-serif">Zip</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3> 
          <INPUT maxLength=100 size=30 id="zip" name="zip">
          </FONT></TD>
      </TR>
      <TR> 
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              contact
<!------------------------------------------------------------------------------------------------------------------------------------ -->
        <TD width=193 height=45> <DIV align=center><font color="#000000" size="3" face="Verdana, Arial, Helvetica, sans-serif">Contact</font></DIV></TD>
        <TD width=307><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3> 
          <INPUT maxLength=100 size=30 id="contact" name="contact">
          </FONT></TD>
      </TR>
      <TR> 
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
        <TD width=193 height=45> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3>Email</FONT></DIV></TD>
        <TD width=307><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3> 
          <INPUT maxLength=40 size=30 id="email" name="email">
          </FONT></TD>
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
          </FONT></TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              USERID
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD height=43> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3>Jersey </FONT></DIV></TD>
        <TD><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3> 
          <input maxlength=3 size=5 id="jersey" name="jersey">
          </FONT></TD>
      </TR>
      <TR> 
        <TD height=42> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3>Userid</FONT></DIV></TD>
        <TD><INPUT maxLength=20 size=30 id="uid" name="uid">
          <FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3>&nbsp; </FONT></TD>
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
          </FONT></TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							         PASSWORD
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD colSpan=2> <P><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2><STRONG>* Mandatory Fields</STRONG></FONT></P>
          <P><STRONG><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2> 
            <INPUT name="adminlevel" type=hidden id="adminlevel" value=4>
            <INPUT type=submit id="option" name="option" value="add">
            <input type="button" name="goback" value="go back" onClick="history.go( -1 );return true;">
            </FONT></STRONG></P></TD>
      </TR>
    </TBODY>
  </TABLE>
  <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							HIDDEN FIELD
<!------------------------------------------------------------------------------------------------------------------------------------ -->
</FORM>
