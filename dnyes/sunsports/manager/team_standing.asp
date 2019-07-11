<script language="JavaScript">
<!--
function gettarg(targ,selObj,restore){ //v3.0
  eval(targ+".location('team_standing.asp?cid="+selObj.options[selObj.selectedIndex].value+"')");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>
<%
cid=request("cid")
if cid="" then cid=210
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
 <table cellpadding=0 cellspacing=0>
  <tbody>
    <tr> 
      <td align="right"><b><font color=#000066 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        size=2>Select the tournament you want to view:</font></b></td>
      <td width=114 align="right"> <select name="cid" class="display_dropdown" id="cid" onchange="gettarg('self',this,0)">
          <option selected value="">SELECT ... </option>
<%
sql="select * from tournament order by cid desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if rs.bof or rs.eof then response.Redirect("error.asp?err=1")
do while not rs.eof
%>
          <%=rs("cid")%> 
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
</table>
<table width="96%" border="0" align="center" cellpadding="2" cellspacing="2">
  <tr> 
    <td> 
      <!-- ========================================================================================================		  

													standing ============star

 ======================================================================================================== -->
      <table width="96%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#660000">
        <tr bgcolor="#FFFFFF"> 
          <td height="25" colspan="9" rowspan="2"><b><%=cname%> Standing / (Tables)</b></td>
        </tr>
        <tr bgcolor="#FFFFFF"> </tr>
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
        </tr>
<%
for i=0 to ubound(ateam)-1
sssql="select * from standing where tid="&ateam(i)
Set ssrs=Server.CreateObject("ADODB.RecordSet") 
ssrs.Open sssql,conn,1,1
if not ( ssrs.eof or ssrs.bof ) then
%>
        <tr> 
          <td align="center" bgcolor="#FFFFFF"><%=ssrs("tname")%></td>
          <td align="center" bgcolor="#FFFFFF"><%=ssrs("played")%></td>
          <td align="center" bgcolor="#FFFFFF"><%=ssrs("win")%></td>
          <td align="center" bgcolor="#FFFFFF"><%=ssrs("draw")%></td>
          <td align="center" bgcolor="#FFFFFF"><%=ssrs("lose")%></td>
          <td align="center" bgcolor="#FFFFFF"><%=ssrs("scored")%></td>
          <td align="center" bgcolor="#FFFFFF"><%=ssrs("conceded")%></td>
          <td align="center" bgcolor="#FFFFFF"><%=ssrs("gd")%></td>
          <td align="center" bgcolor="#FFFFFF"><%=ssrs("points")%></td>
        </tr>
        <%
end if
ssrs.close
set ssrs=nothing
next
%>
        <tr> 
          <td align="center" bgcolor="#FFFFFF">Racing Carnegies FC</td>
          <td align="center" bgcolor="#FFFFFF">24</td>
          <td align="center" bgcolor="#FFFFFF">21</td>
          <td align="center" bgcolor="#FFFFFF">2</td>
          <td align="center" bgcolor="#FFFFFF">1</td>
          <td align="center" bgcolor="#FFFFFF">98</td>
          <td align="center" bgcolor="#FFFFFF">30</td>
          <td align="center" bgcolor="#FFFFFF">68</td>
          <td align="center" bgcolor="#FFFFFF">66</td>
        </tr>
      </table>
      <!-- ========================================================================================================		  

													standing ============stop

 ======================================================================================================== -->
    </td>
  </tr>
</table>
