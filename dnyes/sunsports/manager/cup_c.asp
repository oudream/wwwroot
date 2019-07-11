<SCRIPT LANGUAGE="JavaScript">
<!--

function checkform()
{
	if (zform.zname.value.length == 0) {
		alert(" Enter the name and type of competition. ");
		zform.zname.focus();
		return false;
		}
	if(! isletter(zform.zname.value)) {
		alert(' Sorry ! the name and type of competition makes up by "_" and "0~9" adn "a~z" and "A~Z" and blank space');
		zform.zname.focus();
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
zname=trim(request("zname"))
if zname<>"" then
zxlsurl=request("zxlsurl")
if zxlsurl="" then zxlsurl=" "
zmemo=request("zmemo")
if zmemo="" then zmemo=" "
zmemo=replace(zmemo,"'","¡¯")
sql="insert into tournament (name,type,xlsurl,memo,inserttime) values ('"&zname&"','cup','"&zxlsurl&"','"&zmemo&"','"&now&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Create tournament is complete. ');</SCRIPT>")
end if
%>

<form id="zform" action="cup_c.asp" onSubmit="return checkform();" method="post">
  <table width="100%" border="0" cellspacing="2" cellpadding="2">
    <tr>
      <td width="1%">&nbsp;</td>
      <td colspan="2">&nbsp;</td>
      <td width="4%">&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td colspan="2"><H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>League Wizard </FONT></H2>
        <P><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066 size=2><STRONG>Step 1:</STRONG></FONT> <FONT 
                  face="Verdana, Arial, Helvetica, sans-serif" color=#000000 
                  size=3>Enter the name and type of competition</FONT></P></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="10%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Name :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td width="85%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="zname"  name="zname" maxlength=100 size=30>
        <font color="#000000">*</font> </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="10%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>URL :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td width="85%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="zxlsurl"  name="zxlsurl" maxlength=100 size=30>
        </font>( Cup's Result )</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Memo :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>
        <textarea name="zmemo" cols="50" rows="10" id="zmemo"></textarea>
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td colspan="2"><input type=submit value=Submit name=submit></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td colspan="2"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066 size=2><STRONG>* Mandatory Fields</STRONG></FONT> 
      </td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
