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
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file="adminconn.asp" -->
<%
if request("cid")="" then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Please select the tournament you want to edit at first ! ');</SCRIPT>")
response.End()
end if
%>
<%
if trim(request("option"))="edit" then
cname=trim(request("cname"))
cstatus=request("cstatus")
cmemo=request("cmemo")
if cmemo="" then cmemo=" "
xlsurl=trim(request("xlsurl"))
if xlsurl="" then xlsurl=" "
if cname<>"" then
cmemo=replace(cmemo,"'","''")
xlsurl=replace(xlsurl,"'","''")
cname=replace(cname,"'","''")
sql="update tournament set name='"&cname&"',status='"&cstatus&"',xlsurl='"&xlsurl&"',memo='"&cmemo&"' where cid=" & request("cid")
conn.Execute(sql)
response.Redirect(request("zurl")&"?option=edittournamentissucc")
end if
end if
%>
<%
zurl="league_v.asp"
sql="select * from tournament where cid=" & request("cid")
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
%> 
<form id="form_tournament" action="tournament_e.asp" onSubmit="return checkform();" method="post">
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>Tournament Edit Wizard </FONT></H2>
        <P><FONT 
                  face="Verdana, Arial, Helvetica, sans-serif" color=#000000 
                  >Enter the name and type of competition</FONT></P>
                        </td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Name</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
      <INPUT id="cname"  name="cname" maxLength=100 0 value="<%=rs("name")%>">
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
                  color=#000000 >Present 
<%
if rs("status")="present" or isnull(rs("status")) then present_radio="checked"
if rs("status")="past" then past_radio="checked"
%>
        <INPUT type=radio id="cstatus" name="cstatus" value="present" <%=present_radio%>>
        Past </FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
        <input type=radio id="cstatus" name="cstatus" value="past" <%=past_radio%>>
        </FONT></td>
    <td>&nbsp;</td>
  </tr>
<%
if rs("type")="cup" then
zurl="cup_v.asp"
%>
  <tr>
    <td>&nbsp;</td>
      <td>URL &nbsp;&nbsp;<FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >
        <INPUT id="xlsurl" name="xlsurl" maxLength=100 0 value="<%=rs("xlsurl")%>">
        </FONT> ( Result of competition )</td>
    <td>&nbsp;</td>
  </tr>
<%
end if
%>
  <tr>
    <td>&nbsp;</td>
      <td><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Name</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
        <font color="#000000">
        <textarea name="cmemo" id="cmemo" cols="50" rows="10"><%=rs("memo")%></textarea>
        *</font> </FONT></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>
        <input type="hidden" id="cid" name="cid" value=<%=rs("cid")%>> 
        <input type="hidden" id="zurl" name="zurl" value=<%=zurl%>> 
        <input type="submit" id="option" name="option" value="edit"> 
        <input type="button" name="goback" value="go back" onClick="history.go( -1 );return true;"> 
	</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
      <td height="50">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
<%
rs.close
set rs=nothing
%>
</table>
</form>
<br>
<br>
</body>
</html>
