<script language="JavaScript">

function checkform()
{
	if (form_administrator.ename.value.length == 0) {
		alert("Please enter a valid Name. ");
		form_administrator.ename.focus();
		return false;
		}
	if(! isLetter(form_administrator.ename.value)) {
		alert("Please enter a valid Name. ");
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

<!--#include file="adminconn.asp" -->
<%
eid=request("eid")
if eid="" then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Please select the events you want to edit at first ! ');</SCRIPT>")
response.End()
end if
%>
<%
if trim(request("option"))="del" then
conn.execute "delete * from events where eid=" & eid
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Del events is complete. ');</SCRIPT>")
response.Redirect("events_v.asp")
response.End()
end if
%>
<%
if request("option")="edit" then
ename=trim(request("ename"))
ename=replace(ename,"'","''")
photo=trim(request("photo"))
memo=request("memo")
memo=replace(memo,"'","''")
if memo="" then memo=" "
if ename<>"" then
sql="update events set name='"&ename&"',photo='"&photo&"',memo='"&memo&"' where eid=" & eid
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit events is complete. ');</SCRIPT>")
end if
end if
%>
<H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>Edit NewsEvents</FONT></H2>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<%
sql="select * from events where eid=" & eid
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
%> 
<FORM action="events_e.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
  <TABLE height=307 cellSpacing=0 cellPadding=0 width=500 border=0>
      <TR>
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              NAME
<!------------------------------------------------------------------------------------------------------------------------------------ -->
        <TD width=104 height=45> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3>Name</FONT></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3>
          <INPUT maxLength=100 size=30 id="ename" name="ename" value="<%=rs("name")%>">
          *</FONT></TD>
      </TR>
      <TR>
        <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							              photo
<!------------------------------------------------------------------------------------------------------------------------------------ -->
        <TD width=104 height=45> <DIV align=center><font color="#000000" size="3" face="Verdana, Arial, Helvetica, sans-serif">Photo 
            URL</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3> 
          <INPUT maxLength=100 size=30 id="photo" name="photo" value="<%=rs("photo")%>">
          * </FONT></TD>
      </TR>
      <!---------------------------------------------------------------------------------------------------------------------------------------
<!                     							         memo
<!------------------------------------------------------------------------------------------------------------------------------------ -->
      <TR> 
        <TD height=32> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3>Memo</FONT></DIV></TD>
        <TD><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 size=3> 
          <textarea name="memo" cols="50" rows="10" id="memo"><%=rs("memo")%></textarea>
          </FONT></TD>
      </TR>
      <TR> 
        <TD colSpan=2> <P><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2><STRONG>* Mandatory eventss</STRONG></FONT></P>
          <P><STRONG><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2>
        <input type="hidden" id="eid" name="eid" value=<%=eid%>> 
        <input type="submit" id="option" name="option" value="edit"> 
        <input type="submit" id="option" name="option" value="del" onClick="return confirm(' Are you sure to del the NewsEvents ?')"> 
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