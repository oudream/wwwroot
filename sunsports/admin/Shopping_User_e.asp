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
	if( form_administrator.email.value.length != 0 ){
	if(! isMail(form_administrator.email.value)) {
		alert('Please enter a valid EMAIL address');
		form_administrator.email.focus();
		return false;
		}}
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
pid =request("pid")

if request("option")="edit" then
uname=trim(request("uname"))
uname=replace(uname,"'","''")
tel=trim(request("tel"))
tel=replace(tel,"'","''")
zip=trim(request("zip"))
zip=replace(zip,"'","''")
contact=trim(request("contact"))
contact=replace(contact,"'","''")
email=trim(request("email"))
email=replace(email,"'","''")
uid=trim(request("uid"))
pwd=trim(request("pwd"))



usql="select * from player where pid<>"&pid&"and uid='"&uid&"'"
set urs=server.createobject("ADODB.Recordset")
urs.open usql,conn,1,1
if not ( urs.eof or urs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' The Userid have existed , please enter other Userid . ');</SCRIPT>")
urs.close
set urs=nothing
else
urs.close
set urs=nothing





if tel="" then tel=" "
if zip="" then zip=" "
if contact="" then contact=" "
if email="" then email=" "
if uname<>"" then
conn.Execute("update player set name='"&uname&"',tel='"&tel&"',zip='"&zip&"',contact='"&contact&"',email='"&email&"',uid='"&uid&"',pwd='"&pwd&"' where pid=" & pid)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit user is complete . ');</SCRIPT>")
end if
end if
end if
%>
<%
sql="select * from player where pid=" & pid
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
%>
<H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>Edit Shopping User</FONT></H2>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<FORM action="Shopping_User_e.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
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
        <TD width=104 height=45> <DIV align=center><font color="#000000" size="3" face="Verdana, Arial, Helvetica, sans-serif">Telphone</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3> 
          <INPUT maxLength=100 size=30 id="tel" name="tel" value="<%=rs("tel")%>">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=104 height=45> <DIV align=center><font color="#000000" size="3" face="Verdana, Arial, Helvetica, sans-serif">Zip</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3> 
          <INPUT maxLength=100 size=30 id="zip" name="zip" value="<%=rs("zip")%>">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=104 height=45> <DIV align=center><font color="#000000" size="3" face="Verdana, Arial, Helvetica, sans-serif">Contact</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3> 
          <INPUT maxLength=100 size=30 id="contact" name="contact" value="<%=rs("contact")%>">
          </FONT></TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              email
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD width=104 height=45> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3>Email</FONT></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3> 
          <INPUT maxLength=40 size=30 id="email" name="email" value="<%=rs("email")%>">
          </FONT></TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							             NRIC
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              USERID
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD height=45> <DIV align=center><FONT 
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
        <TD height=45> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3>Password</FONT></DIV></TD>
        <TD><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3> 
          <INPUT maxLength=20 size=30 id="pwd" name="pwd" value="<%=rs("pwd")%>">
          * </FONT></TD>
      </TR>
      <TR> 
        <TD height=32 colspan="2"> <DIV align=center></DIV>
          <FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3>&nbsp; </FONT></TD>
      </TR>
      <TR> 
        <TD colSpan=2> <P><STRONG><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2>
            <INPUT type=submit id="option" name="option" value="edit">
            &nbsp;&nbsp; 
            <INPUT type="button" id="goback" name="goback" value="go back" onClick="history.go( -1 );">
            </FONT></STRONG></P>
          </TD>
      </TR>
    </TBODY>
  </TABLE>
  <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							HIDDEN FIELD
<!------------------------------------------------------------------------------------------------------------------------------------ -->
  			<INPUT name="pid" type=hidden id="pid" value=<%=pid%>>
  			<INPUT name="page" type=hidden id="page" value=<%=page%>>
</FORM>
<%
rs.close
set rs=nothing
%>
