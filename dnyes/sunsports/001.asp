<!--#include file="conn/conn.asp" -->
<table width="96%" border="0" align="center" cellpadding="2" cellspacing="2">
  <tr> 
    <td> 
      <!-- ========================================================================================================		  

													hot shot  ============star

 ======================================================================================================== -->
<%
ssssql="select all * from scorer where pid=3"
Set sssrs=Server.CreateObject("ADODB.RecordSet") 
sssrs.Open ssssql,conn,1,1
if not ( sssrs.eof or sssrs.bof ) then
for i=0 to sssrs.recordcount-1
zzz=zzz+clng(sssrs("score"))
sssrs.movenext
next
end if
sssrs.close
set sssrs=nothing
%>
      <table width="96%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#660000">
        <tr bgcolor="#FFFFFF"> 
          <td height="25" colspan="4"><b><%=cname%> Hot Shots</b></td>
        </tr>
        <tr> 
          <td width="190" height="20" align="center" bgcolor="#FFFFFF"><b><b>Team 
            Name</b></b></td>
          <td width="45" align="center" bgcolor="#FFFFFF"><b>Jersey</b></td>
          <td align="center" bgcolor="#FFFFFF"><b>Name</b></td>
          <td width="90" align="center" bgcolor="#FFFFFF"><b>Goals</b></td>
        </tr>
        <tr> 
          <td align="center" bgcolor="#FFFFFF"> 

<br><br>
<%=zzz%>
<br><br>
          </td>
          <td align="center" bgcolor="#FFFFFF">24</td>
          <td align="center" bgcolor="#FFFFFF">12</td>
          <td align="center" bgcolor="#FFFFFF">2</td>
        </tr>
      </table>
      <!-- ========================================================================================================		  

													hot shot  ============stop

 ======================================================================================================== -->
    </td>
  </tr>
</table>
