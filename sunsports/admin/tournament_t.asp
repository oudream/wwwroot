<script language="JavaScript">
<!--
function gettarg(targ,selObj,restore){ //v3.0
  eval(targ+".location='tournament_t.asp?cid="+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
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
if trim(request("option"))="del" then
conn.execute "delete * from ct where cid=" & request("cid") & " and tid =" & request("tid")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Del the team from the tournament is complete. ');</SCRIPT>")
end if
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td height="10">&nbsp;</td>
        </tr>
        <tr> 
          <td height="30" align="center" valign="middle" class="header">Tournament Edit/Del</td>
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
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td> 
<!-- ========================================================================================================		  

													fixture or result ============star

 ======================================================================================================== -->
            <!-- --------------------------------------------------------------------------------------------------
													assign to "cid"
---------------------------------------------------------------------------------------------------- -->
<%
cid=request("cid")
if cid="" then cid=109
csql="select * from tournament where cid="&cid
Set crs=Server.CreateObject("ADODB.RecordSet")
crs.Open csql,conn,1,1
if crs.eof or crs.bof then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' No league you want in the databse ! ');</SCRIPT>")
response.End()
crs.close
set crs=nothing
end if
cname=crs("name")
crs.close
set crs=nothing
%>
<%
dim ateam()
tsql="select tid from ct where cid="&cid
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
if not ( trs.eof or trs.bof ) then
redim ateam(trs.recordcount-1)
for i=0 to ubound(ateam)
ateam(i)=trs("tid")
trs.movenext
next
trs.close
set trs=nothing
else
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' There are not team in The League(cup) ! ');</SCRIPT>")
%>
            
            <table width="96%" border="0" align="center" cellpadding="2" cellspacing="2">
              <tr>
                <td align="center"><strong><a href="tournament_a.asp?cid=<%=cid%>">ADD 
                  TEAM TO THE TOURNAMENT</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="#" onClick="history.go( -1 );">Go Back 
                 </a></strong></td>
              </tr>
            </table> 
<%
response.End()
end if
%>
<!-- ========================================================================================================		  

													fixture or result ============stop

 ======================================================================================================== -->
<!-- --------------------------------------------------------------------------------------------------
													assign to "team"
---------------------------------------------------------------------------------------------------- -->
            <table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
              <tr> 
                <td height="20" colspan="4" align="center" bgcolor="#FFFFFF"><table cellpadding=0 cellspacing=0>
                    <tbody>
                      <tr> 
                        <td align="right"><b><font color=#000066 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        size=2>Select the tournament you want to view:</font></b></td>
                        <td width=114 align="right"> <select name="cid" class="display_dropdown" id="cid" onchange="gettarg('self',this,0)">
                            <option selected value="">SELECT ... </option>
                            <%
sql="select * from tournament order by cid" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if rs.bof or rs.eof then response.Redirect("error.asp?err=1")
do while not rs.eof
%>
                            <option value=<%=rs("cid")%>><%=rs("name")%> </option>
                            <%
rs.movenext
loop
rs.close
set rs=nothing
%>
                          </select> </td>
                      </tr>
                    </tbody>
                  </table></td>
              </tr>
              <tr> 
                <td width="47" height="20" align="center" bgcolor="#FFFFFF">S/No</td>
                <td width="168" align="center" bgcolor="#FFFFFF">Team Name</td>
                <td align="center" bgcolor="#FFFFFF">Description</td>
                <td width="308" align="center" bgcolor="#FFFFFF">Option</td>
              </tr>
              <%
tsql="select * from team order by name"
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
%>
              <%
if trs.eof or trs.bof then
response.Write("<tr><td rowspan=5 align=center bgcolor=#FFFFFF>NO DATA</td></tr></table>")
response.End()
trs.close
set trs=nothing
else
for j=0 to trs.recordcount-1
for i=0 to ubound(ateam)
if ateam(i)=trs("tid") then
%>
              <tr> 
                <td align="center" bgcolor="#FFFFFF"><%=trs("tid")%></td>
                <td align="center" bgcolor="#FFFFFF"><%=trs("name")%></td>
                <td align="center" bgcolor="#FFFFFF"> <%
if len(trs("description"))>32 then
response.Write(left(trs("description"),30)&"..")
else
response.Write(trs("description"))
end if
%> 				&nbsp;</td>
                <td align="center" bgcolor="#FFFFFF"><a href="tournament_t.asp?option=del&cid=<%=cid%>&tid=<%=trs("tid")%>" onClick="return confirm(' Are you sure del the team from the tournament ? ')">Del 
                  this team from the tournament</a></td>
              </tr>
              <%
end if
next
trs.movenext
next
end if
%>
              <tr> 
                <td height="40" colspan="4" align="center" bgcolor="#FFFFFF"><strong><a href="tournament_a.asp?cid=<%=cid%>">Choose and add other team in this tournament</a></strong></td>
              </tr>
              <%
trs.close
set trs=nothing
%>
            </table></td>
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
