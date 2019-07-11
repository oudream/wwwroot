<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file="conn/conn.asp" -->
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr> 
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td height="10">&nbsp;</td>
        </tr>
        <tr> 
          <td height="30" align="center" valign="middle" class="header">Email View</td>
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
sql="select top 10 * from announcements order by id desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
do while not rs.eof
%>
      <table width="100%" bordercolor="#CCCCCC" border="1" cellpadding="2" cellspacing="0">
        <tr> 
          <td width="73" align="center">NO.</td>
          <td width="792"><%=rs("id")%></td>
          <td width="102" align="center">&nbsp;</td>
        </tr>
        <tr> 
          <td height="19" align="center">User Name</td>
          <td height="19"><%=rs("name")%></td>
          <td align="center">&nbsp;</td>
        </tr>
        <tr> 
          <td height="19" align="center">Insert Time</td>
          <td height="19" align="left"><%=rs("inserttime")%></td>
          <td align="center">&nbsp;</td>
        </tr>
        <tr> 
          <td height="19" align="center">CONTENT</td>
          <td height="19" align="left" valign="top"> <%
if len(rs("content"))>100 then
response.Write(left(rs("content"),98)&"..")
else
response.Write(rs("content"))
end if
%> </td>
          <td align="center">&nbsp;</td>
        </tr>
        <tr> 
          <td height="19" align="center">&nbsp;</td>
          <td height="19" align="center">&nbsp;</td>
          <td align="center">&nbsp;</td>
        </tr>
      </table>
        <%
rs.movenext
loop
rs.close
set rs=nothing
%>
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
