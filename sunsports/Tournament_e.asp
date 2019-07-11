<SCRIPT LANGUAGE="JavaScript">
<!--
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
<%
if request("cid")="" then
response.Redirect("error.asp?err=001")
response.End()
end if
%>
<!--#include file="Top.asp" -->
<%
cname=trim(request("cname"))
cstatus=request("cstatus")
xlsurl=trim(request("xlsurl"))
if xlsurl="" then xlsurl=" "
if cname<>"" then
sql="update tournament set name='"&cname&"',status='"&cstatus&"',xlsurl='"&xlsurl&"' where cid=" & request("cid")
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit tournament is complete. ');</SCRIPT>")
end if
%>
<%
sql="select * from tournament where cid=" & request("cid")
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
%> 
<form id="form_tournament" action="Tournament_e.asp" onSubmit="return checkform();" method="post">
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
      <INPUT id="cname"  name="cname" maxLength=100 size=30 value="<%=rs("name")%>">
      <font color="#000000">*</font> </FONT></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
      <td>Status</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
      <td><FONT face="Verdana, Arial, Helvetica, sans-serif" 
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
<%
if rs("type")="cup" then
%>
  <tr>
    <td>&nbsp;</td>
      <td>URL &nbsp;&nbsp;<FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>
        <INPUT id=xlsurl name=xlsurl maxLength=100 size=30 value="<%=rs("xlsurl")%>">
        </FONT> ( Result of competition )</td>
    <td>&nbsp;</td>
  </tr>
<%
end if
%>
  <tr>
    <td>&nbsp;</td>
    <td>
        <input type="hidden" id="cid" name="cid" value=<%=rs("cid")%>> 
		<input type=submit value=Submit name=submit>
	</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066 size=2><STRONG>* Mandatory Fields</STRONG></FONT> 
    </td>
    <td>&nbsp;</td>
  </tr>
<%
rs.close
set rs=nothing
%>
</table>
</form>
