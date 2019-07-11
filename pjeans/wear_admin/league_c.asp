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
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<!--#include file="adminconn.asp" -->
<%
zname=trim(request("zname"))
if zname<>"" then
zmemo=request("zmemo")
if zmemo="" then zmemo=" "
zmemo=replace(zmemo,"'","''")
zname=replace(zname,"'","''")
sql="insert into tournament (name,type,memo,inserttime) values ('"&zname&"','league','"&zmemo&"','"&now&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Create tournament is complete. ');</SCRIPT>")
end if
%>

<form id="zform" action="league_c.asp" onSubmit="return checkform();" method="post">
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
        <P> <FONT 
                  face="Verdana, Arial, Helvetica, sans-serif" color=#000000 
                  >Enter the name and type of competition</FONT></P></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="10%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Name :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >&nbsp; </FONT></td>
      <td width="85%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff > 
        <input id="zname"  name="zname" maxlength=100 0>
        <font color="#ff0000">*</font> </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 >Memo :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >&nbsp; </FONT></td>
      <td><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff >
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
      <td height="50" colspan="2">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
<br>
<br>
</body>
</html>
