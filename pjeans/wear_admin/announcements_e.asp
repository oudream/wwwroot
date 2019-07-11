<script language="JavaScript">

function checkform()
{
	if (form_administrator.inserttime.value.length == 0) {
		alert("Please enter a valid Sent Time. ");
		form_administrator.inserttime.focus();
		return false;
		}
	if (form_administrator.content.value.length == 0) {
		alert("Please enter a valid content. ");
		form_administrator.content.focus();
		return false;
		}
	if(! isLetter(form_administrator.content.value)) {
		alert("Please enter a valid content. ");
		form_administrator.content.focus();
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
aid=request("id")
if aid="" then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Please select a announcements you want to edit at first ! ');</SCRIPT>")
response.End()
end if
%>
<%
if trim(request("option"))="del" then
response.Redirect("announcements_v.asp?option=del&id="&request("id"))
end if
%>
<%
if request("option")="edit" then
inserttime=trim(request("inserttime"))
inserttime=replace(inserttime,"'","''")
content=trim(request("content"))
content=replace(content,"'","''")
sql="update announcements set content='"&content&"',inserttime='"&inserttime&"' where id=" & aid
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit the announcement is complete . ');</SCRIPT>")
end if
%>
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr> 
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td height="10">&nbsp;</td>
        </tr>
        <tr> 
          <td height="30" align="center" valign="middle" class="header_a"><strong>Email View</strong></td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td> 
<!-- ========================================================================================================		  

													cup table ============star

 ======================================================================================================== -->
<%
sql="select * from announcements where id="&aid&" order by id desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if not ( rs.bof or rs.eof ) then
%>
<FORM action="announcements_e.asp" method=post name="form_administrator" id="form_administrator" onSubmit="return checkform();">
        <table width="98%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
          <tr> 
            <td><strong>SENT TIME:</strong></td>
            <td> <input type="text" id="inserttime" name="inserttime" maxlength="100" size="30" value="<%=rs("inserttime")%>"> 
            </td>
          </tr>
          <tr> 
            <td><strong>CONTENT:</strong></td>
            <td><textarea name="content" id="content" cols="50" rows="10"><%=rs("content")%></textarea></td>
          </tr>
          <tr align="center"> 
            <td colspan="2">
        <input type="hidden" id="id" name="id" value=<%=aid%>> 
        <input type="submit" id="option" name="option" value="edit"> 
        <input type="submit" id="option" name="option" value="del" onClick="return confirm(' Are you sure to del the tournament ?')"> 
        <input type="button" name="goback" value="go back" onClick="history.go( -1 );return true;"> 
			  </td>
          </tr>
        </table>
      </form>
      <!-- ========================================================================================================		  

													cup table ============stop

 ======================================================================================================== -->
    </td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<%
end if
rs.close
set rs=nothing
%>
</body>
</html>
