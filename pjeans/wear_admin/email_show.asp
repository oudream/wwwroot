<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file="adminconn.asp" -->
<%
eid=request("id")
if eid="" then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Please select a email you want to edit at first ! ');</SCRIPT>")
response.End()
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
sql="select * from mail where id="&eid&" order by id desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if not ( rs.bof or rs.eof ) then
%>
      <table width="98%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
        <tr> 
          <td><strong>SENT TIME:</strong></td>
          <td><%=rs("inserttime")%></td>
        </tr>
        <tr> 
          <td width="116" height="25"><strong>NAME:</strong></td>
          <td><%=rs("name")%> </td>
        </tr>
        <tr> 
          <td><strong>EMAIL:</strong></td>
          <td><%=rs("email")%></td>
        </tr>
        <tr> 
          <td><strong>SUBJECT:</strong></td>
          <td><%=rs("subject")%></td>
        </tr>
        <tr> 
          <td><strong>CONTENT:</strong></td>
          <td><%=rs("mailtext")%> </td>
        </tr>
        <tr align="center"> 
          <td colspan="2"><a href="mailto:<%=rs("email")%>"><strong>REPLY THE EMAIL</strong></a></td>
        </tr>
      </table>
      <!-- ========================================================================================================		  

													cup table ============stop

 ======================================================================================================== -->
    </td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td align="center"><a href="javascript:history.back()"><br>
      <br>
      BACK</a></td>
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
