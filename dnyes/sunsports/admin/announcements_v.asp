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
conn.execute "delete * from announcements where id=" & request("id")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Del announcement is complete. ');</SCRIPT>")
end if
%>
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr> 
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td align="center" valign="middle" class="header"><H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>Announcements Edit/Delete</FONT></H2></td>
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
          <td width="70" align="center">Userid</td>
          <td width="84" align="center">User Name</td>
          <td width="118" align="center">Insert Time</td>
          <td width="335" align="center">CONTENT</td>
          <td colspan="2" align="center">Edit/Del</td>
        </tr>
        <%
sql="select * from announcements order by id desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if rs.bof or rs.eof then response.Redirect("error.asp?err=1")
do while not rs.eof
%>
        <tr> 
          <td height="19" align="center"><%=rs("id")%>&nbsp;</td>
          <td height="19" align="center">
		  <%
		  	psql="select * from player where pid="&rs("pid") 
			Set prs=Server.CreateObject("ADODB.RecordSet") 
			prs.Open psql,conn,1,1 
			if not prs.eof then
			response.Write(prs("uid"))
			end if
			prs.close
			set prs=nothing
		  %>&nbsp;
		  </td>
          <td align="center"><%=rs("name")%>&nbsp;</td>
          <td align="center"><%=rs("inserttime")%>&nbsp;</td>
          <td align="center"> <%
if len(rs("content"))>100 then
response.Write(left(rs("content"),98)&"..")
else
response.Write(rs("content"))
end if
%> &nbsp;</td>
          <td width="39" align="center"><a href="announcements_e.asp?id=<%=rs("id")%>">Edit</a></td>
          <td width="40" align="center"><a href="announcements_v.asp?id=<%=rs("id")%>&option=del"% onClick="return confirm(' Are you sure to del the Email ?')">Del</a></td>
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
