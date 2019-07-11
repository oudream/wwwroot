<SCRIPT LANGUAGE="JavaScript">
<!--

function checkform()
{
	if (zform.uname.value.length == 0) {
		alert(" Enter the name and type of competition. ");
		zform.uname.focus();
		return false;
		}
	if(! isletter(zform.uname.value)) {
		alert(' Sorry ! the name and type of competition makes up by "_" and "0~9" adn "a~z" and "A~Z" and blank space');
		zform.uname.focus();
		return false;
	}
	if(! isMail(zform.email.value)) {
		alert('Please enter a valid EMAIL address');
		zform.email.focus();
		return false;
		}
	if(! isNumber(zform.jersey.value)) {
		alert('Use integers only for "Jersey No." field');
		zform.jersey.focus();
		return false;
		}
	if (zform.uid.value.length > 1) {
	if(! isLetter(zform.uid.value)) {
		alert('The "Userid" field makes up by "_" and "0~9" adn "a~z" and "A~Z" and blank space');
		zform.uid.focus();
		return false;
		}
		}
	if (zform.pwd.value.length > 1) {
	if(! isLetter(zform.pwd.value)) {
		alert('The "Password" field makes up by "_" and "0~9" adn "a~z" and "A~Z" and blank space');
		zform.pwd.focus();
		return false;
		}
		}
	return true;
}

function isletter(name)
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
	if(name.length==0)
		return true;
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

//-->

</SCRIPT>

<!--#include file="adminconn.asp" -->
<%
ztid=request("ztid")
if request("option")="Add" then
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

sql="insert into player (name,uid,pwd,tel,zip,tid,contact,email,nric,jersey,default_score,default_yellow,default_red,ptype,inserttime,adminlevel) values ('"&uname&"','"&uid&"','"&pwd&"','"&tel&"','"&zip&"',"&ztid&",'"&contact&"','"&email&"','"&nric&"',"&jersey&","&default_score&","&default_yellow&","&default_red&",'"&"p','"&now&"',0)"
conn.Execute(sql)
response.Redirect("team_t.asp?option=createplayersucc&tid="&ztid)
end if
end if
%>

<form id="zform" action="player_c.asp" onSubmit="return checkform();" method="post">
		<input type="hidden" name="ztid" id="ztid" value=<%=ztid%>>
  <table width="100%" border="0" cellspacing="2" cellpadding="2">
    <tr> 
      <td width="1%">&nbsp;</td>
      <td colspan="2">&nbsp;</td>
      <td width="4%">&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td height="50" colspan="2"><H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>Create Player Wizard </FONT></H2></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="22%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Name :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td width="73%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="uname"  name="uname" maxlength=50 size=30>
        <font color="#000000">*</font> </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="22%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Telphone :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td width="73%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="tel"  name="tel" maxlength=20 size=20>
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="22%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Zip :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td width="73%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="zip"  name="zip" maxlength=10 size=10>
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="22%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Contact :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td width="73%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="contact"  name="contact" maxlength=100 size=50>
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="22%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Email :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td width="73%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="email"  name="email" maxlength=30 size=30>
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="22%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Nric/Passort No. :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td width="73%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="nric"  name="nric" maxlength=10 size=30>
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="22%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Jersey :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td width="73%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="jersey"  name="jersey" maxlength=3 size=5>
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="22%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Userid:</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td width="73%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="uid"  name="uid" maxlength=20 size=30>
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="22%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Password :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td width="73%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="pwd"  name="pwd" maxlength=20 size=30>
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td colspan="2"> <input type="submit" value="Add" name="option" id="option"> 
        <input type="button" name="goback" value="go back" onClick="history.go( -1 );return true;"> 
      </td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td height="50" colspan="2">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
