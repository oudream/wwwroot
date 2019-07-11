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
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

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
                  color=#000066>Tournament Edit Wizard </FONT></H2>
        <P> <FONT 
                  face="Verdana, Arial, Helvetica, sans-serif" color=#000000 
                  >Enter the name and type of competition</FONT></P></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="10%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Name :</FONT></td>
      <td width="85%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
        <input id="zname"  name="zname" maxlength=100 0 <%=rs("name")%>>
        <font color="#000000">*</font> </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="10%">Status :</td>
      <td width="85%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >&nbsp; </font><FONT face="Verdana, Arial, Helvetica, sans-serif" 
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
    <tr> 
      <td>&nbsp;</td>
      <td><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Note :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >&nbsp; </FONT></td>
      <td><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >
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
      <td height="50" colspan="2">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
<br>
<br>
</body>
</html>
