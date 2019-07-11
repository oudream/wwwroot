<script language="JavaScript">

function checkform()
{
	if (form_administrator.ename.value.length == 0) {
		alert("Please enter a valid Name. ");
		form_administrator.ename.focus();
		return false;
		}
	if(! isLetter(form_administrator.ename.value)) {
		alert('The "Name" events makes up by "_" and "0~9" adn "a~z" and "A~Z" and blank space');
		form_administrator.ename.focus();
		return false;
		}
	if (form_administrator.photo.value.length == 0) {
		alert("Please enter a valid Photo Address. ");
		form_administrator.photo.focus();
		return false;
		}
	if (! isEnglish(form_administrator.photo.value)) {
		alert("Please enter a valid Photo Address. ");
		form_administrator.photo.focus();
		return false;
		}
	if (! isEnglish(form_administrator.memo.value)) {
		alert("Please enter a valid Memo. ");
		form_administrator.memo.focus();
		return false;
		}
	return true;
}

function isLetter(name)
{
	for(i = 0; i < name.length; i++) {
		if(name.charCodeAt(i) > 128)
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

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file="adminconn.asp" -->
<%
if request("option")="add" then
ename=trim(request("ename"))
ename=replace(ename,"'","''")
photo=trim(request("photo"))
memo=request("memo")
memo=replace(memo,"'","''")
if memo="" then memo=" "
if ename<>"" then
sql="insert into events (name,photo,memo,inserttime) values ('"&ename&"','"&photo&"','"&memo&"','"&now()&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Create events is complete. ');</SCRIPT>")
end if
url="events_v.asp?option=addsucc"
response.redirect url
end if
%>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center"><H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066><br>
        Create NewsEvents<br>
        </FONT></H2>
</td>
  </tr>
</table>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<FORM action="events_c.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
  <TABLE width=98% border=0 align="center" cellPadding=0 cellSpacing=0>
      <TR>
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              NAME
<!------------------------------------------------------------------------------------------------------------------------------------ -->
        <TD width=104 height=45> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >Title</FONT></DIV></TD>
        <TD width=493><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >
          <INPUT maxLength=100 0 id="ename" name="ename">
          *</FONT></TD>
      </TR>
      <TR>
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              photo
<!------------------------------------------------------------------------------------------------------------------------------------ -->
        <TD width=104 height=45> <DIV align=center><font color="#000000"  face="Verdana, Arial, Helvetica, sans-serif">Photo 
            URL</font></DIV></TD>
        
      <TD width=493><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
        <INPUT maxLength=100 0 id="photo" name="photo" value="events_images/">
        * ( Photo in &quot;events_images/ &quot; )</FONT></TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							         memo
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD height=32> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >Detail</FONT></DIV></TD>
        <TD><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <textarea name="memo" cols="50" rows="10" id="memo"></textarea>
          </FONT></TD>
      </TR>
      <TR> 
        <TD colSpan=2> <P><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2><STRONG>* Mandatory events</STRONG></FONT></P>
          <P><STRONG><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2>
            <INPUT type=submit id="option" name="option" value="add">
            </FONT></STRONG></P></TD>
      </TR>
  </TABLE>
  <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							HIDDEN events
<!------------------------------------------------------------------------------------------------------------------------------------ -->
</FORM>
<P>&nbsp;</P>
