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
if request("option")="editasstmanagerissucc" then response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit asst.manager is complete. ');</SCRIPT>")
if request("option")="editmanagerissucc" then response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit manager is complete. ');</SCRIPT>")
tid=session("manager_tid")
if tid="" then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Please select the team you want to add standing at first ! ');</SCRIPT>")
response.End()
end if
%>
<%
if trim(request("option"))="del" then
conn.execute "delete * from player where pid=" & request("pid")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Del the player from the team is complete. ');</SCRIPT>")
end if
%>
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr> 
    <td>&nbsp;</td>
    <td><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td height="30" class="header"><h2><font color="#000066"><br>
              Team's Player Edit/Del<br>
              </font></h2></td>
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
          <td> 
<!-- ========================================================================================================		  

													fixture or result ============star

 ======================================================================================================== -->
            <!-- --------------------------------------------------------------------------------------------------
													assign to "tid"
---------------------------------------------------------------------------------------------------- -->
<!-- ========================================================================================================		  

													fixture or result ============stop

 ======================================================================================================== -->
<!-- --------------------------------------------------------------------------------------------------
													assign to "team"
---------------------------------------------------------------------------------------------------- -->
            <table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
              <tr> 
                <td height="20" colspan="5" align="center" bgcolor="#FFFFFF">&nbsp;</td>
              </tr>
              <tr> 
                <td width="71" height="20" align="center" bgcolor="#FFFFFF">S/No</td>
                <td width="199" align="center" bgcolor="#FFFFFF">Team Name</td>
                <td width="449" align="center" bgcolor="#FFFFFF">Member Type</td>
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
i=0
do while not prs.eof
i=i+1
%>
              <tr> 
                <td align="center" bgcolor="#FFFFFF"><%=i%></td>
                <td align="center" bgcolor="#FFFFFF"><%=prs("name")%></td>
                <td align="center" bgcolor="#FFFFFF">&nbsp;
<%
if prs("ptype")="m" then response.Write("Manager")
if prs("ptype")="a" then response.Write("Asst. Manager")
if prs("ptype")="p" then response.Write("Player")
%>				
				</td>
                <td width="91" align="center" bgcolor="#FFFFFF">
<%
if prs("adminlevel")=2 then
%>
	<a href="manager_e.asp?pid=<%=prs("pid")%>"> EDIT </a> 
<%
elseif prs("adminlevel")=5 then
%>
	<a href="manager_asst_e.asp?pid=<%=prs("pid")%>"> EDIT </a> 
<%
elseif prs("ptype")="p" then
%>
				<a href="player_e.asp?option=edit&pid=<%=prs("pid")%>" >Edit</a>
<%
end if
%>
				</td>
                <td width="130" align="center" bgcolor="#FFFFFF">
<%
if not ( prs("ptype")="m" or prs("ptype")="a" ) then 
%>				
				<a href="team_t.asp?option=del&pid=<%=prs("pid")%>" onClick="return confirm(' Are you sure del the player from the team ? ')" title="Del this player from the team">Delete</a>
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
                  A PLAYER</a></strong></td>
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
