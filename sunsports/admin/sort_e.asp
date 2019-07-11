<script language="JavaScript">
function checkform()
{
	if (form_administrator.name.value.length == 0) {
		alert("Please enter a valid name. ");
		form_administrator.name.focus();
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
sort_id=request("sort_id")
if sort_id="" then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Please select a sort you want to edit at first ! ');</SCRIPT>")
response.End()
end if
%>
<%
if request("option")="edit" then
zname=trim(request("name"))
zname=replace(zname,"'","''")
sql="update sort set name='"&zname&"' where sort_id=" & sort_id
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit the sort is complete . ');</SCRIPT>")
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
          <td height="30" align="center" valign="middle" class="header_a"><strong>Sort 
            Edit </strong></td>
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
sql="select * from sort where sort_id="&sort_id&" order by sort_id desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if not ( rs.bof or rs.eof ) then
%>
<FORM action="sort_e.asp" method=post name="form_administrator" id="form_administrator" onSubmit="return checkform();">
        <table width="98%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
          <tr> 
            <td width="116"><strong> NAME :</strong></td>
            <td><input type="text" id="name" name="name" maxlength="50" size="30" value="<%=rs("name")%>"></td>
          </tr>
          <tr align="center"> 
            <td colspan="2"> <input type="hidden" id="sort_id" name="sort_id" value=<%=sort_id%>> 
              <input type="submit" id="option" name="option" value="edit">&nbsp;&nbsp;&nbsp;
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
