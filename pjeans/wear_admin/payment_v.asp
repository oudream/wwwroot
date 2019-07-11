<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file="adminconn.asp" -->
<%
if trim(request("option"))="del" then
conn.execute "delete * from payment where payment_id=" & request("payment_id")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Del payment is complete. ');</SCRIPT>")
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
          <td height="30" align="center" valign="middle" class="header">payment View</td>
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
      <table width="100%" bordercolor="#CCCCCC" border="1" cellpadding="2" cellspacing="0">
        <tr> 
          <td width="50" align="center">NO.</td>
          <td width="84" align="center">payment</td>
          <td width="84" align="center">Content</td>
          <td colspan="2" align="center">Operation</td>
        </tr>
        <%
sql="select * from payment order by payment_id desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if rs.bof or rs.eof then response.Redirect("error.asp?err=1")
do while not rs.eof
%>
        <tr> 
          <td height="19" align="center"><%=rs("payment_id")%>&nbsp;</td>
          <td align="center"><%=rs("name")%>&nbsp;</td>
          <td align="center">
		  <%
		  if len(rs("content"))>102 then
		  response.Write(left(rs("content"),100)&"..")
		  else
		  response.Write(rs("content"))
		  end if
		  %>
		  &nbsp;</td>
          <td width="39" align="center"><a href="payment_e.asp?payment_id=<%=rs("payment_id")%>">Edit</a></td>
          <td width="40" align="center"><a href="payment_v.asp?payment_id=<%=rs("payment_id")%>&option=del"% onClick="return confirm(' Are you sure to del the Email ?')">Del</a></td>
        </tr>
        <%
rs.movenext
loop
rs.close
set rs=nothing
%>
      </table>
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
</body>
</html>
