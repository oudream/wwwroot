<SCRIPT LANGUAGE="JavaScript">
<!--

function checkform()
{
	if (form_z.zname.value.length == 0) {
		alert(" Enter the name and type of competition. ");
		form_z.zname.focus();
		return false;
		}
	if(! isletter(form_z.zname.value)) {
		alert(' Sorry ! the name and type of competition makes up by "_" and "0~9" adn "a~z" and "A~Z" and blank space');
		form_z.zname.focus();
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

<!--#include file="../conn/conn.asp" -->
<%
zname=trim(request("zname"))
cstatus=request("cstatus")
note=request("note")
if zname<>"" then
sql="insert into tournament (name,status,note) values ('"&zname&"','"&cstatus&"','"&note&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit tournament is complete. ');</SCRIPT>")
end if
%>
<%
sql="select * from tournament where cid=" & request("cid")
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
%> 

<form id="form_z" action="league_c.asp" onSubmit="return checkform();" method="post">
  <table width="100%" border="0" cellspacing="2" cellpadding="2">
    <tr>
      <td width="1%">&nbsp;</td>
      <td colspan="2">&nbsp;</td>
      <td width="4%">&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td colspan="2"><H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>Tournament Wizard </FONT></H2>
        <P><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066 size=2><STRONG>Step 1:</STRONG></FONT> <FONT 
                  face="Verdana, Arial, Helvetica, sans-serif" color=#000000 
                  size=3>Enter the name and type of competition</FONT></P></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="10%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Name :</FONT></td>
      <td width="85%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="zname"  name="zname" maxlength=100 size=30 <%=rs("name")%>>
        <font color="#000000">*</font> </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="10%">Status :</td>
      <td width="85%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </font><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Present 
<%
if rs("status")="present" or isnull(rs("status")) then present_radio="checked"
if rs("status")="past" then past_radio="checked"
%>
        <INPUT type=radio id="cstatus" name="cstatus" value="present" <%=present_radio%>>
        Past </FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input type=radio id="cstatus" name="cstatus" value="past" <%=past_radio%>>
        </FONT></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Note :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>
        <textarea name="note" cols="50" rows="10" id="note"></textarea>
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
