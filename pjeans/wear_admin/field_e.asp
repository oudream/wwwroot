<script language="JavaScript">

function checkform()
{
	if (form_administrator.fname.value.length == 0) {
		alert("Please enter a valid Name. ");
		form_administrator.fname.focus();
		return false;
		}
	if(! isLetter(form_administrator.fname.value)) {
		alert("Please enter a valid Name. ");
		form_administrator.fname.focus();
		return false;
		}
	if (form_administrator.location.value.length == 0) {
		alert("Please enter a valid Location. ");
		form_administrator.location.focus();
		return false;
		}
	if(! isEnglish(form_administrator.location.value)) {
		alert("Please enter a valid Location. ");
		form_administrator.location.focus();
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
fid=request("fid")
if fid="" then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Please select the field you want to edit at first ! ');</SCRIPT>")
response.End()
end if
%>
<%
if trim(request("option"))="del" then
conn.execute "delete * from field where fid=" & fid
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Del field is complete. ');</SCRIPT>")
response.Redirect("field_v.asp")
response.End()
end if
%>
<%
if request("option")="edit" then
fname=trim(request("fname"))
fname=replace(fname,"'","''")
location=trim(request("location"))
photo=trim(request("photo"))
tfnight=request("tfnight")
memo=request("memo")
memo=replace(memo,"'","''")
if memo="" then memo=" "
if fname<>"" then
sql="update field set name='"&fname&"',location='"&location&"',photo='"&photo&"',tfnight='"&tfnight&"',memo='"&memo&"' where fid=" & fid
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit field is complete. ');</SCRIPT>")
end if
end if
%>

<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center"><H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066><br>
        Edit Field<br>
        </FONT></H2>
</td>
  </tr>
</table>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<%
sql="select * from field where fid=" & fid
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
%> 
<FORM action="field_e.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
  <TABLE width=98% border=0 align="center" cellPadding=0 cellSpacing=0>
      <TR>
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              NAME
<!------------------------------------------------------------------------------------------------------------------------------------ -->
        <TD width=104 height=45> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >Name</FONT></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >
          <INPUT maxLength=100 0 id="fname" name="fname" value="<%=rs("name")%>">
          *</FONT></TD>
      </TR>
      <TR>
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              location
<!------------------------------------------------------------------------------------------------------------------------------------ -->
        <TD width=104 height=45> <DIV align=center><font color="#000000"  face="Verdana, Arial, Helvetica, sans-serif">Location</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 0 id="location" name="location" value="<%=rs("name")%>">
          * </FONT></TD>
      </TR>
      <TR>
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              photo
<!------------------------------------------------------------------------------------------------------------------------------------ -->
        <TD width=104 height=45> <DIV align=center><font color="#000000"  face="Verdana, Arial, Helvetica, sans-serif">Photo 
            URL</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 0 id="photo" name="photo" value="<%=rs("photo")%>">
          * </FONT></TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							             NRIC
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD height=43> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >Night(Y/N)</FONT></DIV></TD>
        <TD><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
<%
checkedstr=rs("tfnight")
if checkedstr="yes" then acheck="checked"
if checkedstr="no" or checkedstr="" then bcheck="checked"

%>
          <input type="radio" name="tfnight" id="tfnight" value="yes" <%=acheck%>>
          YES 
          <input type="radio" name="tfnight" id="tfnight" value="no" <%=bcheck%>>
          NO </FONT></TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							         memo
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD height=32> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >Memo</FONT></DIV></TD>
        <TD><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <textarea name="memo" cols="50" rows="10" id="memo"><%=rs("memo")%></textarea>
          </FONT></TD>
      </TR>
      <TR> 
        <TD colSpan=2> <P><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2><STRONG>* Mandatory Fields</STRONG></FONT></P>
          <P><STRONG><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2>
        <input type="hidden" id="fid" name="fid" value=<%=fid%>> 
        <input type="submit" id="option" name="option" value="edit"> 
        <input type="submit" id="option" name="option" value="del" onClick="return confirm(' Are you sure to del the tournament ?')"> 
        <input type="button" name="goback" value="go back" onClick="history.go( -1 );return true;"> 
            </FONT></STRONG></P></TD>
      </TR>
  </TABLE>
  <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							HIDDEN FIELD
<!------------------------------------------------------------------------------------------------------------------------------------ -->
</FORM>
<%
rs.close
set rs=nothing
%>