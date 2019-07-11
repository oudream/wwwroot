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
	if(form_administrator.email.value.length > 0){
	if(! isMail(form_administrator.email.value)) {
		alert('Please enter a valid EMAIL address');
		form_administrator.email.focus();
		return false;
		}}
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
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="8" topmargin="8" marginwidth="0" marginheight="0">
<!--#include file="adminconn.asp" -->
<%
adminlevel=request("adminlevel")
if adminlevel=1 then
adminstr="administrator"
elseif adminlevel=3 then
adminstr="visitor"
else
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' error ! ');</SCRIPT>")
response.End()
end if
if request("option")="add" then
uname=trim(request("uname"))
uname=replace(uname,"'","''")
uid=trim(request("uid"))
pwd=trim(request("pwd"))
if uid<>"" and uname<>"" and pwd<>"" then
usql="select * from player where uid='"&uid&"'"
set urs=server.createobject("ADODB.Recordset")
urs.open usql,conn,1,1
if not ( urs.eof or urs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' The Userid have existed , please enter other Userid . ');</SCRIPT>")
urs.close
set urs=nothing
else
urs.close
set urs=nothing
contact=trim(request("contact"))
contact=replace(contact,"'","''")
tel=trim(request("tel"))
tel=replace(tel,"'","''")
zip=trim(request("zip"))
zip=replace(zip,"'","''")
email=trim(request("email"))
email=replace(email,"'","''")
nric=trim(request("nric"))
if tel="" then tel=" "
if zip="" then zip=" "
if contact="" then contact=" "
if email="" then email=" "
if nric="" then nric=" "
sql="insert into player (name,tel,zip,contact,email,uid,pwd,nric,adminlevel,inserttime) values ('"&uname&"','"&tel&"','"&zip&"','"&contact&"','"&email&"','"&uid&"','"&pwd&"','"&nric&"',"&adminlevel&",'"&now()&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Create administrator is complete. ');</SCRIPT>")
end if
end if
end if
%>
<H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>Create <%=adminstr%></FONT></H2>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<FORM action="user_c.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
  <TABLE height=307 cellSpacing=0 cellPadding=0 width=500 
                  border=0>
    <TBODY>
      <TR> 
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              NAME
<!------------------------------------------------------------------------------------------------------------------------------------ -->
        <TD width=193 height=45> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >Name</FONT></DIV></TD>
        <TD width=307><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="uname" name="uname">
          *</FONT></TD>
      </TR>
      <TR> 
        <TD width=104 height=45> <DIV align=center><font color="#000000"  face="Verdana, Arial, Helvetica, sans-serif">Telphone</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="tel" name="tel">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=104 height=45> <DIV align=center><font color="#000000"  face="Verdana, Arial, Helvetica, sans-serif">Zip</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="zip" name="zip">
          </FONT></TD>
      </TR>
      <TR> 
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              contact
<!------------------------------------------------------------------------------------------------------------------------------------ -->
        <TD width=193 height=45> <DIV align=center><font color="#000000"  face="Verdana, Arial, Helvetica, sans-serif">Contact</font></DIV></TD>
        <TD width=307><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="contact" name="contact">
          </FONT></TD>
      </TR>
      <TR> 
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
        <TD width=193 height=45> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >Email</FONT></DIV></TD>
        <TD width=307><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=40 size=30 id="email" name="email">
          </FONT></TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							             NRIC
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD height=43> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >Nric/Passort No.</FONT></DIV></TD>
        <TD><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=10 size=30 id="nric" name="nric">
          </FONT></TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              USERID
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD height=42> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >Userid</FONT></DIV></TD>
        <TD><INPUT maxLength=20 size=30 id="uid" name="uid"> <FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >*</FONT></TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							         PASSWORD
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD height=32> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >Password</FONT></DIV></TD>
        <TD><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=20 size=30 id="pwd" name="pwd">
          * </FONT></TD>
      </TR>
      <TR> 
        <TD height="50" colSpan=2> <P><STRONG><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2>
<INPUT type=submit id="option" name="option" value="add">
            </FONT></STRONG></P>
          </TD>
      </TR>
    </TBODY>
  </TABLE>
  <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							HIDDEN FIELD
<!------------------------------------------------------------------------------------------------------------------------------------ -->
  			<INPUT name="adminlevel" type=hidden id="adminlevel" value=<%=adminlevel%>>
  			<INPUT name="page" type=hidden id="page" value=<%=page%>>
</FORM>
<br>
<br>
</body>
</html>
