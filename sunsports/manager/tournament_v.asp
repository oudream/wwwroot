<html>
<head>
<title>Sunsport Adminstrator</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body>
<!--#include file="adminconn.asp" -->
<%
if trim(request("option"))="del" then
conn.execute "delete * from tournament where cid=" & request("cid")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Del tournament is complete. ');</SCRIPT>")
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
          <td height="30" align="center" valign="middle"><strong>Tournament Edit/Del</strong></td>
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
          <td width="6%">NO.</td>
          <td width="9%">NAME</td>
          <td>TYPE</td>
          <td>STATUS</td>
          <td>URL</td>
          <td>MEMO</td>
          <td colspan="2">Edit/Del</td>
        </tr>
<%
sql="select * from tournament order by cid desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if rs.bof or rs.eof then response.Redirect("error.asp?err=1")
do while not rs.eof
%>
        <tr> 
          <td width="6%" height="19"><a href="League.asp?cid=<%=rs("cid")%>" title="<%=rs("name")%>"> 
            <%=rs("cid")%></a>&nbsp;</td>
          <td width="9%"><a href="League.asp?cid=<%=rs("cid")%>" title="<%=rs("name")%>"><%=rs("name")%></a>&nbsp;</td>
          <td width="3%"><a href="League.asp?cid=<%=rs("cid")%>" title="<%=rs("name")%>"><%=rs("type")%></a>&nbsp;</td>
          <td width="6%"><a href="League.asp?cid=<%=rs("cid")%>" title="<%=rs("name")%>"><%=rs("status")%></a>&nbsp;</td>
          <td width="13%"><a href="League.asp?cid=<%=rs("cid")%>" title="<%=rs("name")%>"><%=rs("xlsurl")%></a>&nbsp;</td>
          <td width="43%">
<%
if len(rs("memo"))>16 then
response.Write(left(rs("memo"),15)&"..")
else
response.Write(rs("memo"))
end if
%>
		  &nbsp;</td>
          <td width="4%"><a href="tournament_e.asp?cid=<%=rs("cid")%>" title="<%=rs("name")%>">Edit</a></td>
          <td width="16%"><a href="tournament_v.asp?cid=<%=rs("cid")%>&option=del"% onClick="return confirm(' Are you sure to del the tournament ?')">Del</a></td>
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
