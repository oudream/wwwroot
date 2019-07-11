<script language="JavaScript">

function checkform()
{
	if (form_administrator.colorname.value.length == 0) {
		alert("Please enter a valid Color Name. ");
		form_administrator.colorname.focus();
		return false;
		}
	if(! isEnglish(form_administrator.colorname.value)) {
		alert("Please enter a valid Color Name. ");
		form_administrator.colorname.focus();
		return false;
		}
	if (form_administrator.color.value.length == 0) {
		alert("Please enter a valid color . ");
		form_administrator.color.focus();
		return false;
		}
	if (! isLetter(form_administrator.color.value)) {
		alert(' Sorry ! the Color makes up by "0~9" adn "a~f" and "A~F" ');
		form_administrator.color.focus();
		return false;
		}
	return true;
}

function isLetter(name)
{
var letter="0123456789abcdefABCDEF"
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
<body text="#000000" leftmargin="8" topmargin="8" marginwidth="0" marginheight="0">

<!--#include file="adminconn.asp" -->
<%
id=request("id")
if id="" then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Please select the Color you want to edit at first ! ');</SCRIPT>")
response.End()
end if
%>
<%
if request("option")="edit" then
colorname=trim(request("colorname"))
colorname=replace(colorname,"'","''")
color=trim(request("color"))
sql="update color set colorname='"&colorname&"',color='"&color&"' where id=" & id
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit Color is complete. ');</SCRIPT>")
end if
%>
<H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>Edit Color</FONT></H2>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<%
sql="select * from color where id=" & id
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
%> 
<FORM action="color_e.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
  <TABLE height=165 cellSpacing=0 cellPadding=0 width=597 border=0>
    <TR> 
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              NAME
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TD width=149 height=45> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >Color name :</FONT></DIV></TD>
      <TD width=448><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
        <INPUT maxLength=100 size=30 id="colorname" name="colorname" value="<%=rs("colorname")%>">
        *</FONT></TD>
    </TR>
    <TR> 
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              color
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TD width=149 height=45> <DIV align=center><font color="#000000"  face="Verdana, Arial, Helvetica, sans-serif">Color 
          : </font></DIV></TD>
      <TD width=448><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
        <INPUT maxLength=100 size=30 id="color" name="color" value="<%=rs("color")%>">
        * ( Like &quot; FF0000&quot; )</FONT></TD>
    </TR>
    <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							         memo
<!------------------------------------------------------------------------------------------------------------------------------------ -->
    <TR> 
      <TD height="75" colSpan=2> <P><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2></FONT></P>
        <P><STRONG><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2> 
          <input type="hidden" id="id" name="id" value=<%=id%>> 
          <INPUT type=submit id="option" name="option" value="edit">
          <input type="button" name="goback" value="go back" onClick="history.go( -1 );return true;"> 
          </FONT></STRONG></P></TD>
    </TR>
  </TABLE>
  <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							HIDDEN events
<!------------------------------------------------------------------------------------------------------------------------------------ -->
</FORM>
<%
rs.close
set rs=nothing
%>
<br>
<br>
</body>
</html>
