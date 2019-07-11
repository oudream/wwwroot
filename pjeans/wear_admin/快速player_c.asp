<SCRIPT LANGUAGE="JavaScript">
<!--

function checkform()
{
	if (zform.uname.value.length == 0) {
		alert(" Enter the name and type of competition. ");
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
	if(! isNumber(zform.default_score.value)) {
		alert('Use integers only for "Score" field');
		zform.default_score.focus();
		return false;
		}
	if(! isNumber(zform.default_yellow.value)) {
		alert('Use integers only for "Yellow Card" field');
		zform.default_yellow.focus();
		return false;
		}
	if(! isNumber(zform.default_red.value)) {
		alert('Use integers only for "Red Card" field');
		zform.default_red.focus();
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
if nric="" then nric=0
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
      <td height="50" colspan="2"><h2><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>Create Player Wizard </font></h2>
        <table width="100%" border="1" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="215" align="center">Nric /Passport </td>
            <td width="227" align="center">Name</td>
            <td width="116" align="center">Contact</td>
            <td width="107" align="center">Email</td>
            <td width="69" align="center">Jersey No</td>
            <td width="45" align="center">Goals</td>
            <td width="45" align="center" bgcolor="#FFFF00">Yellow</td>
            <td width="44" align="center" bgcolor="#FF0000">Red</td>
          </tr>
          <tr> 
            <td align="center"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
              <input id="nric3"  name="nric" maxlength=10 size=20>
              </font></td>
            <td align="center"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
              <input id="uname2"  name="uname" maxlength=50 size=20>
              </font></td>
            <td align="center"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
              <input id="tel3"  name="tel" maxlength=20 size=15>
              </font></td>
            <td align="center"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
              <input id="email3"  name="email" maxlength=30 size=10>
              </font></td>
            <td align="center"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
              <input id="jersey3"  name="jersey" maxlength=3 size=5>
              </font></td>
            <td align="center"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
              <input id="default_score4"  name="default_score" maxlength=3 size=5>
              </font></td>
            <td align="center"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
              <input id="default_yellow4"  name="default_yellow" maxlength=3 size=5>
              </font></td>
            <td align="center"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
              <input id="default_red3"  name="default_red" maxlength=3 size=5>
              </font></td>
          </tr>
        </table></td>
      <td valign="bottom">
	  <input type="submit" value="Add" name="option" id="option2">
      </td>
    </tr>
  </table>
  <p>&nbsp;</p>
  <p>TEMP!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!TEMP!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!TEMP!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!<br>
    PLAYER_C.ASP HAS BEEN BACKUPED!</p>
</form>
