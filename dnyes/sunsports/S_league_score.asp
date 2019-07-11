<!--#include file="conn/conn.asp" -->
<table width="96%" border="0" align="center" cellpadding="2" cellspacing="2">
  <tr> 
    <td> 
      <!-- ========================================================================================================		  

													hot shot  ============star

 ======================================================================================================== -->
<%
dim aplayer()
psql="select distinct pid from scorer where tid=0 "
for i=0 to 1
psql=psql&" or tid=" & i+1
next
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1 
if not ( prs.eof or prs.bof ) then
redim aplayer(prs.recordcount-1)
for i=0 to ubound(aplayer)
aplayer(i)=prs("pid")
prs.movenext
next
prs.close
set prs=nothing
else
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' There are not player in The League(cup) ! ');</SCRIPT>")
response.End()
prs.close
set prs=nothing
end if
%>
<%
dim dd()
dim amaxscorer()
dim amaxpid()
redim amaxscorer(ubound(aplayer))
redim amaxpid(ubound(aplayer))


for i=0 to ubound(aplayer)


ssssql="select * from scorer where pid="&aplayer(i)
Set sssrs=Server.CreateObject("ADODB.RecordSet") 
sssrs.Open ssssql,conn,1,1
if not ( sssrs.eof or sssrs.bof ) then
redim dd(sssrs.recordcount-1)
cc=0
for j=0 to sssrs.recordcount-1
dd(cc)=sssrs("score")
amaxscorer(i)=amaxscorer(i)+clng(sssrs("score"))
sssrs.movenext
cc=cc+1
next
amaxpid(i)=aplayer(i)
end if
sssrs.close
set sssrs=nothing


next


j=ubound(amaxscorer)
for i=0 to ubound(amaxscorer)
for k=0 to j
if k=j then exit for
if amaxscorer(k)>amaxscorer(k+1) then
tmpmaxscorer=amaxscorer(k) 
amaxscorer(k)=amaxscorer(k+1)
amaxscorer(k+1)=tmpmaxscorer
tmpmaxpid=amaxpid(k) 
amaxpid(k)=amaxpid(k+1)
amaxpid(k+1)=tmpmaxpid
end if
next
j=j-1
next
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
            <%
for i=0 to ubound(amaxscorer)
response.Write(cstr(amaxscorer(i))+"<br>")
next
%>
<br><br>
            <%
for i=0 to ubound(dd)
response.Write(cstr(dd(i))+"<br>")
next
%>

<br><br>
<%=cc%>
<br><br>
<%
for i=0 to ubound(aplayer)
response.Write(cstr(aplayer(i))+"<br>")
next
%>
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
