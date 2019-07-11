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
conn.execute "delete * from sort where sort_id=" & request("sort_id")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Del sort is complete. ');</SCRIPT>")
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
          <td height="30" align="center" valign="middle" class="header">Sort View</td>
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
          <td width="84" align="center">Sort</td>
          <td colspan="2" align="center">Operation</td>
        </tr>
        <%
sql="select * from sort order by sort_id desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if rs.bof or rs.eof then response.Redirect("error.asp?err=1")
do while not rs.eof
%>
        <tr> 
          <td height="19" align="center"><%=rs("sort_id")%>&nbsp;</td>
          <td align="center"><%=rs("name")%>&nbsp;</td>
          <td width="39" align="center"><a href="sort_e.asp?sort_id=<%=rs("sort_id")%>">Edit</a></td>
          <td width="40" align="center"><a href="sort_v.asp?sort_id=<%=rs("sort_id")%>&option=del"% onClick="return confirm(' Are you sure to del the Email ?')">Del</a></td>
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
