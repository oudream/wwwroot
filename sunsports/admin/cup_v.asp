<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file="adminconn.asp" -->
<%
if request("option")="edittournamentissucc" then response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit tournament is complete. ');</SCRIPT>")
if trim(request("option"))="del" then
conn.Execute("delete * from tournament where cid=" & request("cid"))
conn.Execute("delete * from ct where cid=" & request("cid"))
conn.Execute("delete * from match where cid=" & request("cid"))
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Del tournament is complete. ');</SCRIPT>")
end if
%>
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr> 
    <td><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td height="10">&nbsp;</td>
        </tr>
        <tr> 
          <td height="30" align="center" valign="middle" class="header">Tournament Edit/Del</td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td> 
      <!-- ========================================================================================================		  

													cup table ============star

 ======================================================================================================== -->
      <table width="100%" border="1" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
        <tr bgcolor="#FFFFFF"> 
          <td width="40" align="center">NO.</td>
          <td width="200" align="center">NAME</td>
          <td width="50" align="center">TYPE</td>
          <td width="50" align="center">STATUS</td>
          <td width="150" align="center">URL</td>
          <td align="center">MEMO</td>
          <td colspan="2" align="center">Edit/Del</td>
        </tr>
        <%
sql="select * from tournament where type='cup' order by cid" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if rs.bof or rs.eof then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' There are not cup in tournament . ');</SCRIPT>")
response.End()
end if
do while not rs.eof
%>
        <tr bgcolor="#FFFFFF"> 
          <td width="45" height="19" align="center"><%=rs("cid")%>&nbsp;</td>
          <td align="center"><%=rs("name")%>&nbsp;</td>
          <td align="center"><%=rs("type")%>&nbsp;</td>
          <td align="center"><%=rs("status")%>&nbsp;</td>
          <td align="center"><%=rs("xlsurl")%>&nbsp;</td>
          <td> <%
if len(rs("memo"))>16 then
response.Write(left(rs("memo"),15)&"..")
else
response.Write(rs("memo"))
end if
%> &nbsp;</td>
          <td width="30" align="center"><a href="tournament_e.asp?cid=<%=rs("cid")%>" title="<%=rs("name")%>">Edit</a></td>
          <td width="30" align="center"><a href="cup_v.asp?cid=<%=rs("cid")%>&option=del"% onClick="return confirm(' Are you sure to del the tournament ?')">Del</a></td>
        </tr>
        <%
rs.movenext
loop
rs.close
set rs=nothing
%>
        <tr bgcolor="#FFFFFF"> 
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
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
