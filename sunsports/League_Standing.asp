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
