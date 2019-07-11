<!--#include file="Top_sun.asp"-->
<table width="776" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="181" valign="top"> <br><br><br>


      <!-- ========================================================================================================		  

													Tournament table ============star

 ======================================================================================================== -->
      <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
        <tr> 
          <td><a href="League.asp"><b><font color="#1F4A71">LEAGUE MATCHES</font></b></a><font color="#1F4A71">&nbsp;+ 
            +</font></td>
        </tr>
      </table> 
      <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
        <tr> 
          <td><a href="Cup.asp"><b><font color="#1F4A71">CUP MATCHES </font></b></a><font color="#1F4A71">&nbsp;+ 
            +</font></td>
        </tr>
      </table>
      <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
        <tr> 
          <td><a href="Friendlies.asp"><b><font color="#1F4A71">FRIENDLY MATCHES 
            </font></b></a></td>
        </tr>
      </table>
      <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
        <tr> 
          <td><a href="TeamProfile.asp"><b><font color="#1F4A71">TEAMPORFILE</font></b></a></td>
        </tr>
      </table>
      <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
        <%
sql="select * from tournament order by type desc , id_order" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if rs.bof or rs.eof then response.Redirect("error.asp?err=1")
do while not rs.eof
%>
        <tr> 
          <td width="5" valign="top"><a href="teamprofile.asp?cid=<%=rs("cid")%>" title="<%=rs("name")%>" target="_blank"> 
            </a><img src="images/sun_l_sp.gif" width="5" height="10" border="0"></td>
          <td><a href="teamprofile.asp?cid=<%=rs("cid")%>" title="<%=rs("name")%>" class="b"><%=rs("name")%></a></td>
        </tr>
        <%
rs.movenext
loop
rs.close
set rs=nothing
%>
      </table>
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
response.End()
end if
%>
      <!-- ========================================================================================================		  

													Tournament table ============star

 ======================================================================================================== -->
    </td>
    <td width="579" align="center" valign="top"><br>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
		  <td width="20">&nbsp;</td>
          <td>
<!-- ========================================================================================================		  

													fixture or result ============star

 ======================================================================================================== -->
 
 <!-- --------------------------------------------------------------------------------------------------
													assign to "cid"
---------------------------------------------------------------------------------------------------- -->
<!-- ========================================================================================================		  

													fixture or result ============stop

 ======================================================================================================== -->
<!-- --------------------------------------------------------------------------------------------------
													assign to "team"
---------------------------------------------------------------------------------------------------- -->
<%
tsql="select * from team order by tid"
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
%>
		  <table width="100%" border="0" align="center" cellpadding="10" cellspacing="1" bgcolor="#6184A3">
              <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                          onmouseover="this.style.backgroundColor='#BECFDF';"> 
                <td height="20" colspan="2" align="center"><b><%=cname%></b></td>
              </tr>
              <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                          onmouseover="this.style.backgroundColor='#BECFDF';"> 
                <td width="77" height="20" align="center"><strong>S/No</strong></td>
                <td width="484" align="center"><strong>Team Name</strong></td>
              </tr>
              <%
if trs.eof or trs.bof then
response.Write("<tr><td colspan=2 align=center bgcolor=#FFFFFF>NO DATA</td></tr></table>")
response.End()
trs.close
set trs=nothing
else
for j=0 to trs.recordcount-1
for i=0 to ubound(ateam)
if ateam(i)=trs("tid") then
%>
              <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                          onmouseover="this.style.backgroundColor='#BECFDF';"> 
                <td align="center"><%=j+1%></td>
                <td align="center"><a href="ShowTeam.asp?tid=<%=trs("tid")%>&cid=<%=cid%>" target="_blank" class="b"><%=trs("name")%></a></td>
              </tr>
              <%
end if
next
trs.movenext
next
end if
trs.close
set trs=nothing
%>
            </table></td>
        </tr>
      </table>
      <p>&nbsp;</p></td>
  </tr>
</table>
<!--#include file="Copyright.asp"-->