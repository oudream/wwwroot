<html>
<head>
<title>Sunsport Adminstrator</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body>
<!--#include file="adminconn.asp" -->
<%
if trim(request("option"))="del" then
conn.execute "delete * from team where tid=" & request("tid")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Del team is complete. ');</SCRIPT>")
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
          <td height="30" align="center" valign="middle"><strong>Team Edit/Del</strong></td>
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
      <table width="98%" bordercolor="#CCCCCC" border="1" cellpadding="2" cellspacing="0">
        <tr> 
          <td width="4%">NO.</td>
          <td width="10%">NAME</td>
          <td>City</td>
          <td>Create Date</td>
          <td>Manager</td>
          <td>Asst.Manager </td>
          <td colspan="2">Edit/Del</td>
        </tr>
<%
sql="select * from team" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if rs.bof or rs.eof then response.Redirect("error.asp?err=1")
do while not rs.eof
%>
        <tr> 
          <td width="4%" height="19"><a href="League.asp?tid=<%=rs("tid")%>" title="<%=rs("name")%>"> 
            <%=rs("tid")%></a>&nbsp;</td>
          <td width="10%"><a href="League.asp?tid=<%=rs("tid")%>" title="<%=rs("name")%>"><%=rs("name")%></a>&nbsp;</td>
          <td width="11%"><a href="League.asp?tid=<%=rs("tid")%>" title="<%=rs("name")%>"><%=rs("city")%></a>&nbsp;</td>
          <td width="16%"><a href="League.asp?tid=<%=rs("tid")%>" title="<%=rs("name")%>"><%=rs("createdate")%></a>&nbsp;</td>
          <td width="3%">
<%
psql="select * from player where pid=" & rs("mpid")
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1 
if not ( prs.bof or prs.eof ) then 
response.Write(prs("name"))
end if
prs.close
set prs=nothing
%>
		  &nbsp;</td>
          <td width="37%">
<%
psql="select * from player where pid=" & rs("apid")
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1 
if not ( prs.bof or prs.eof ) then 
response.Write(prs("name"))
end if
prs.close
set prs=nothing
%>
            &nbsp;</td>
          <td width="5%"><a href="team_e.asp?tid=<%=rs("tid")%>" title="<%=rs("name")%>">Edit</a></td>
          <td width="14%"><a href="team_v.asp?tid=<%=rs("tid")%>&option=del"% onClick="return confirm(' Are you sure to del the tournament ?')">Del</a></td>
        </tr>
        <%
rs.movenext
loop
rs.close
set rs=nothing
%>
        <tr> 
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
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
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
