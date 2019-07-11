<SCRIPT LANGUAGE="JavaScript">
<!--
function getcup(value)
{
		if (value=="cup"){
		tfcup.style.display="block";}
		if (value=="league"){
		tfcup.style.display="none";}
}
	
function checkform()
{
	if (form_tournament.cname.value.length == 0) {
		alert(" Enter the name and type of competition. ");
		form_tournament.cname.focus();
		return false;
		}
	if(! isletter(form_tournament.cname.value)) {
		alert(' Sorry ! the name and type of competition makes up by "_" and "0~9" adn "a~z" and "A~Z" and blank space');
		form_tournament.cname.focus();
		return false;
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
//-->
</SCRIPT>

<!--#include file="adminconn.asp" -->
<%
cname=trim(request("cname"))
ctype=request("ctype")
xlsurl=trim(request("xlsurl"))
if xlsurl="" then xlsurl=" "
if cname<>"" then
sql="insert into tournament (name,type,xlsurl) values ('"&cname&"','"&ctype&"','"&xlsurl&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Create tournament is complete. ');</SCRIPT>")
end if
%>

<form id="form_tournament" action="tournament_c.asp" onSubmit="return checkform();" method="post">
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>Tournament Wizard </FONT></H2>
      <P><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066 size=2><STRONG>Step 1:</STRONG></FONT> <FONT 
                  face="Verdana, Arial, Helvetica, sans-serif" color=#000000 
                  size=3>Enter the name and type of competition</FONT></P>
                        </td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Name</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
      <INPUT id="cname"  name="cname" maxLength=100 size=30>
      <font color="#000000">*</font> </FONT></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Type</FONT></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>League 
      <INPUT type=radio value=league id="ctype" name="ctype" CHECKED onClick='getcup("league")' >
      Cup </FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
      <INPUT type=radio value=cup id="ctype" name="ctype" onClick='getcup("cup")' >
      </FONT></td>
    <td>&nbsp;</td>
  </tr>
  <tr id="tfcup" style="display:none;" >
    <td>&nbsp;</td>
      <td>URL &nbsp;&nbsp;<FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>
        <INPUT id=xlsurl name=xlsurl maxLength=100 size=30>
        </FONT> ( Result of competition )</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><input type=submit value=Submit name=submit></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066 size=2><STRONG>* Mandatory Fields</STRONG></FONT> 
    </td>
    <td>&nbsp;</td>
  </tr>
</table>
</form>
