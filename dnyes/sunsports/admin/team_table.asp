<script language="JavaScript">
<!--
function gettarg(targ,selObj,restore){ //v3.0
  eval(targ+".location='team_table.asp?cid="+selObj.options[selObj.selectedIndex].value+"'");
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
cid=request("cid")
if cid="" then cid=210
if request("option")="createtableissucc" then response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Create standing is complete. ');</SCRIPT>")
if request("option")="edittableissucc" then response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit standing is complete. ');</SCRIPT>")
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
%>
            <table width="96%" border="0" align="center" cellpadding="2" cellspacing="2">
              <tr>
                <td align="center"><strong><a href="tournament_a.asp?cid=<%=cid%>">ADD 
                  TEAM TO THE TOURNAMENT</a></strong></td>
              </tr>
            </table> 
<%
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
          <td height="30" align="center" valign="middle" class="header">
<%
sql="select name from tournament where cid="&cid
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if not rs.eof then
cname=rs("name")
end if
rs.close
set rs=nothing
%>
		  Tournament <font color="#FF0000"><%=cname%></font> Team's Standing / Table Edit/Del</td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td align="center" valign="middle"><table cellpadding=0 cellspacing=0>
        <tbody>
          <tr> 
            <td align="right"><b><font color=#000066 
                        face="Verdana, Arial, Helvetica, sans-serif">Select the tournament you want to view:</font></b></td>
            <td width=114 align="right"> <select name="cid" class="display_dropdown" id="cid" onChange="gettarg('self',this,0)">
                <option selected value="">SELECT ... </option>
<%
sql="select * from tournament order by cid desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if not ( rs.bof or rs.eof ) then 
do while not rs.eof
%>
                <option value=<%=rs("cid")%>><%=rs("name")%> </option>
                <%
rs.movenext
loop
end if
rs.close
set rs=nothing
%>
              </select> </td>
          </tr>
        </tbody>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><table width="100%" border="0" align="center" cellpadding="2" cellspacing="2">
        <tr> 
          <td> 
            <!-- ========================================================================================================		  

													standing ============star

 ======================================================================================================== -->
            <table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#660000">
              <tr bgcolor="#FFFFFF"> 
                <td height="25" colspan="10"><b><%=cname%> Standing / (Tables)</b></td>
              </tr>
              <tr> 
                <td width="190" height="20" align="center" bgcolor="#FFFFFF"><b>Team 
                  Name</b></td>
                <td width="45" align="center" bgcolor="#FFFFFF"><b>P</b></td>
                <td width="45" align="center" bgcolor="#FFFFFF"><b>W</b></td>
                <td width="45" align="center" bgcolor="#FFFFFF"><b>D</b></td>
                <td width="45" align="center" bgcolor="#FFFFFF"><b>L</b></td>
                <td width="45" align="center" bgcolor="#FFFFFF"><b>F</b></td>
                <td width="45" align="center" bgcolor="#FFFFFF"><b>A</b></td>
                <td width="45" align="center" bgcolor="#FFFFFF"><b>GD</b></td>
                <td width="45" align="center" bgcolor="#FFFFFF"><b>PT</b></td>
                <td width="45" align="center" bgcolor="#FFFFFF"><strong>Option</strong></td>
              </tr>
<%
for i=0 to ubound(ateam)
sssql="select * from standing where tid="&ateam(i)
Set ssrs=Server.CreateObject("ADODB.RecordSet") 
ssrs.Open sssql,conn,1,1
if not ( ssrs.eof or ssrs.bof ) then
%>
              <tr> 
                <td bgcolor="#FFFFFF"><%=ssrs("tname")%></td>
                <td align="center" bgcolor="#FFFFFF"><%=ssrs("played")%></td>
                <td align="center" bgcolor="#FFFFFF"><%=ssrs("win")%></td>
                <td align="center" bgcolor="#FFFFFF"><%=ssrs("draw")%></td>
                <td align="center" bgcolor="#FFFFFF"><%=ssrs("lose")%></td>
                <td align="center" bgcolor="#FFFFFF"><%=ssrs("scored")%></td>
                <td align="center" bgcolor="#FFFFFF"><%=ssrs("conceded")%></td>
                <td align="center" bgcolor="#FFFFFF"><%=ssrs("gd")%></td>
                <td align="center" bgcolor="#FFFFFF"><%=ssrs("points")%></td>
                <td align="center" bgcolor="#FFFFFF"><strong><a href="team_table_e.asp?cid=<%=cid%>&tid=<%=ssrs("tid")%>">EDIT</a></strong></td>
              </tr>
<%
else
tsql="select * from team where tid="&ateam(i)
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1
if not ( trs.eof or trs.bof ) then
%>
              <tr> 
                <td bgcolor="#FFFFFF"><%=trs("name")%></td>
                <td colspan="9" align="center" bgcolor="#FFFFFF">The team  have not standing , you can : <a href="team_table_c.asp?cid=<%=cid%>&tid=<%=trs("tid")%>">click 
                  here to create a new one</a> &nbsp;</td>
              </tr>
              <%
end if
trs.close
set trs=nothing
%>
              <%
end if
ssrs.close
set ssrs=nothing
next
%>
            </table>
            <!-- ========================================================================================================		  

													standing ============stop

 ======================================================================================================== -->
          </td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
</table>
