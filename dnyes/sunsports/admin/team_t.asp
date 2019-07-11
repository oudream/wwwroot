<script language="JavaScript">
<!--
function gettarg(targ,selObj,restore){ //v3.0
  eval(targ+".location='team_t.asp?tid="+selObj.options[selObj.selectedIndex].value+"'");
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
tid=request("tid")
if tid="" then tid=1
if request("option")="editmanagersucc" then response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit manager is complete. ');</SCRIPT>")
if request("option")="editmanagerasstsucc" then response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit Asst.manager is complete. ');</SCRIPT>")
if request("option")="editplayersucc" then response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit Player is complete. ');</SCRIPT>")
if request("option")="createplayersucc" then response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Create player is complete. ');</SCRIPT>")
%>
<%
if trim(request("option"))="del" then
conn.Execute("delete * from player where pid=" & request("pid"))
conn.Execute("delete * from card where pid=" & request("pid"))
conn.Execute("delete * from scorer where pid=" & request("pid"))
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Del the player from the team is complete. ');</SCRIPT>")
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
          <td height="30" align="center" valign="middle" class="header">
<%
tsql="select * from team where tid=" & tid 
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
tname=trs("name")
trs.close
set trs=nothing
%>
		  <font color="#FF0000"><%=tname%></font> Player Edit/Del</td>
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
      <table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td align="center"> 
<strong><a href="player_c.asp?ztid=<%=tid%>">CREATE 
                  APLAYER TO THE TOURNAMENT</a></strong>            <table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
              <tr> 
                <td height="20" colspan="5" align="center" bgcolor="#FFFFFF"><table cellpadding=0 cellspacing=0>
                    <tbody>
                      <tr> 
                        <td align="right"><b><font color=#000066 
                        face="Verdana, Arial, Helvetica, sans-serif">Select the team you want to view:</font></b></td>
                        <td width=114 align="right"> <select name="tid" class="display_dropdown" id="tid" onChange="gettarg('self',this,0)">
                            <option selected value="">SELECT ... </option>
<%
sql="select * from team order by name" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if rs.bof or rs.eof then response.Redirect("error.asp?err=1")
do while not rs.eof
%>
                            <option value=<%=rs("tid")%>><%=rs("name")%> </option>
                            <%
rs.movenext
loop
rs.close
set rs=nothing
%>
                          </select></td>
                      </tr>
                    </tbody>
                  </table></td>
              </tr>
              <tr> 
                <td width="41" height="20" align="center" bgcolor="#FFFFFF">S/No</td>
                <td width="128" align="center" bgcolor="#FFFFFF">Team Name</td>
                <td width="285" align="center" bgcolor="#FFFFFF">Ptype</td>
                <td colspan="2" align="center" bgcolor="#FFFFFF">Option</td>
              </tr>
<%
psql="select * from player where tid=" & tid
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1 
if prs.eof or prs.bof then
response.Write("<tr><td rowspan=5 align=center bgcolor=#FFFFFF>NO DATA</td></tr></table>")
response.End()
prs.close
set prs=nothing
end if
do while not prs.eof
%>
              <tr> 
                <td align="center" bgcolor="#FFFFFF"><%=prs("pid")%>&nbsp;</td>
                <td align="center" bgcolor="#FFFFFF"><%=prs("name")%>&nbsp;</td>
                <td align="center" bgcolor="#FFFFFF">
<%
if prs("ptype")="m" then response.Write("Manager")
if prs("ptype")="a" then response.Write("Asst. Manager")
if prs("ptype")="p" then response.Write("Player")
%>				
				&nbsp;</td>
                <td width="56" align="center" bgcolor="#FFFFFF">
<%
if prs("ptype")="m" then 
%>				
		        <a href="manager_e.asp?zurl=team_t.asp&tid=<%=tid%>&pid=<%=prs("pid")%>">Edit</a>
<%
end if
%>              
<%
if prs("ptype")="a" then 
%>				
        <a href="manager_asst_e.asp?zurl=team_t.asp&tid=<%=tid%>&pid=<%=prs("pid")%>">Edit</a>
<%
end if
%>              
<%
if prs("ptype")="p" then 
%>				
		<a href="player_e.asp?zurl=team_t.asp&tid=<%=tid%>&pid=<%=prs("pid")%>">Edit</a>
<%
end if
%>              
				&nbsp;</td>
                <td width="80" align="center" bgcolor="#FFFFFF">
<%
if not ( prs("ptype")="m" or prs("ptype")="a" ) then 
%>				
				<a href="team_t.asp?option=del&pid=<%=prs("pid")%>" onClick="return confirm(' Are you sure del the player from the team ? ')">Del from the team</a>
<%
end if
%>              
				  &nbsp;</td>
				</tr>
<%
prs.movenext
loop
%>
              <tr> 
                <td colspan="5" align="center" bgcolor="#FFFFFF"><strong><a href="player_c.asp?ztid=<%=tid%>">CREATE 
                  APLAYER TO THE TOURNAMENT</a></strong></td>
              </tr>
<%
prs.close
set prs=nothing
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
