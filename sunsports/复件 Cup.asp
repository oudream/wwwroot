<!--#include file="Top_sun.asp"-->
<table width="776" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="181" valign="top"> <br>
      <table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td> 
            <!-- ========================================================================================================		  

													cup table ============star

 ======================================================================================================== -->
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="League.asp"><b><font color="#1F4A71">LEAGUE MATCHES</font></b></a><font color="#1F4A71">+ 
                  +</font></td>
              </tr>
            </table> 
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="Cup.asp"><b><font color="#1F4A71">CUP MATCHES </font></b></a></td>
              </tr>
            </table> 
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <%
sql="select * from tournament where type='cup' order by id_order" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if rs.bof or rs.eof then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' There are not cup in tournament . ');</SCRIPT>")
response.End()
end if
do while not rs.eof
%>
              <tr> 
                <td width="5" valign="top"><img src="images/sun_l_sp.gif" width="5" height="10" border="0"></td>
                <td><a href="Cup.asp?cid=<%=rs("cid")%>" title="<%=rs("name")%>" class="b"><%=rs("name")%></a></td>
              </tr>
              <%
rs.movenext
loop
rs.close
set rs=nothing
%>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="Friendlies.asp"><b><font color="#1F4A71">FRIENDLY MATCHES 
                  </font></b></a></td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="TeamProfile.asp"><b><font color="#1F4A71">TEAMPORFILE</font></b></a> 
                  <font color="#1F4A71">&nbsp;+ +</font> </td>
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
      
    </td>
    <td width="579" valign="top"><br>
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
		  <td width="20">&nbsp;</td>
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
xlsurl=crs("xlsurl")
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

<!-- --------------------------------------------------------------------------------------------------
													assign to "mweek"
---------------------------------------------------------------------------------------------------- -->
<%
msql="select mweek from match where cid="&cid&" order by mweek desc" 
Set mrs=Server.CreateObject("ADODB.RecordSet") 
mrs.Open msql,conn,1,1 
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#6184A3">
  <tr bgcolor="#FFFFFF"> 
                <td height="25" colspan="11"><b> <%=cname%> -- Fixture and Result &nbsp;&nbsp;<a href="<%=xlsurl%>" target="_blank" onClick="return confirm(' There are not team in The League(cup) ! ');">(click here to see Results map) </a></b></td>
  </tr>
<%
if mrs.eof or mrs.bof then
response.Write("<tr bgcolor=#FFFFFF><td height=25 colspan=11><b>NO DATA</b></td></tr></table>")
response.End()
mrs.close
set mrs=nothing
end if
mweek=request("mweek")
if mweek="" then mweek=mrs("mweek")
mrs.close
set mrs=nothing
%>
    <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                          onmouseover="this.style.backgroundColor='#BECFDF';"> 
    <td width="60" align="center"><b>Date</b></td>
    <td width="35" align="center"><b>Time</b></td>
    <td width="35" align="center"><b>Day</b></td>
    <td width="45" align="center"><b>Home Confirm</b></td>
    <td align="center"><b>Home Team</b></td>
    <td width="13" align="center"><b>CL</b></td>
    <td width="35" align="center"><b>VS</b></td>
    <td width="13"><b>CL</b></td>
    <td align="center"><b>Away Team</b></td>
    <td width="45" align="center"><b>Away Confirm</b></td>
    <td width="91" align="center"><b>Location</b></td>
  </tr>
<%
sql="select * from match where mweek="&mweek&" and cid="&cid&" order by myear,mmonth,mday,mhour,mminute desc" 
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.Open sql,conn,1,1
%>
<%
do while not rs.eof
matchdata=cstr(rs("mmonth"))&"/"&cstr(rs("mday"))&"/"&cstr(rs("myear"))
matchtime=cstr(rs("mhour"))&":"&cstr(rs("mminute"))
%>
    <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                          onmouseover="this.style.backgroundColor='#BECFDF';"> 
    <td height="35" align="center"><%=matchdata%> </td>
    <td align="center"><%=matchtime%></td>
    <td align="center">
<%
if weekday(matchdata,1)=1 then response.Write("SUN")
if weekday(matchdata,1)=2 then response.Write("MON")
if weekday(matchdata,1)=3 then response.Write("TUE")
if weekday(matchdata,1)=4 then response.Write("WED")
if weekday(matchdata,1)=5 then response.Write("THU")
if weekday(matchdata,1)=6 then response.Write("FRI")
if weekday(matchdata,1)=7 then response.Write("SAT")
%>
	</td>
    <td align="center"><%=rs("hconfirm")%></td>
    <td height="35" align="center"><%=rs("htname")%>
	<br></td>
    <td bgcolor=<%=rs("hcolor")%>>&nbsp;</td>
    <td align="center">
	
<% if trim(rs("matched"))="yes" then %>
                  <%
ssql="select * from scorer where mid="&rs("mid")
Set srs=Server.CreateObject("ADODB.RecordSet") 
srs.Open ssql,conn,1,1
if srs.eof or srs.bof then
response.Write("NO DATA")
srs.close
set srs=nothing
else
hsumscorer=0
asumscorer=0
do while not srs.eof
if srs("tid")=rs("htid") then hsumscorer=srs("score")+hsumscorer
if srs("tid")=rs("atid") then asumscorer=srs("score")+asumscorer
srs.movenext
loop
srs.close
set srs=nothing
end if
%>
<%
response.Write(cstr(hsumscorer) & ":" & cstr(asumscorer))
else
response.Write("-")
end if
%>
	</td>
    <td bgcolor=<%=rs("acolor")%>>&nbsp;</td>
    <td align="center"><%=rs("atname")%></td>
    <td align="center"><%=rs("aconfirm")%></td>
    <td align="center"><%=rs("fname")%></td>
  </tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
</table>
		  
<!-- ========================================================================================================		  

													fixture or result ============stop

 ======================================================================================================== -->		  
		  
          </td>
        </tr>
      </table> <br>
      <br>
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
		  <td width="20">&nbsp;</td>
          <td> 
<!-- ========================================================================================================		  

													hot shot  ============star

 ======================================================================================================== -->
<%
dim aplayer()
psql="select distinct pid from scorer where tid=0 "
for i=0 to ubound(ateam)
psql=psql&" or tid="&ateam(i)
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
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' There are not player in The League(cup) of Hot Shot ! ');</SCRIPT>")
response.End()
prs.close
set prs=nothing
end if
%>
<%
dim amaxscorer()
dim amaxpid()
redim amaxscorer(ubound(aplayer))
redim amaxpid(ubound(aplayer))
for i=0 to ubound(aplayer)


ssssql="select * from scorer where pid="&aplayer(i)
Set sssrs=Server.CreateObject("ADODB.RecordSet") 
sssrs.Open ssssql,conn,1,1
if not ( sssrs.eof or sssrs.bof ) then
for j=0 to sssrs.recordcount-1
amaxscorer(i)=amaxscorer(i)+clng(sssrs("score"))
sssrs.movenext
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
            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#6184A3">
              <tr bgcolor="#FFFFFF"> 
                <td height="25" colspan="4"><b><%=cname%> Hot Shots</b></td>
              </tr>
                <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                                      onmouseover="this.style.backgroundColor='#BECFDF';"> 
                <td width="190" height="20" align="center"><b><b>Team 
                  Name</b></b></td>
                <td width="45" align="center"><b>Jersey</b></td>
                <td align="center"><b>Name</b></td>
                <td width="91" align="center"><b>Goals</b></td>
              </tr>

<%
if ubound(aplayer)>19 then 
j=19
else
j=ubound(aplayer)
end if
for i=0 to j
sssssql="select * from scorer where pid="&amaxpid(j-i)
Set ssssrs=Server.CreateObject("ADODB.RecordSet") 
ssssrs.Open sssssql,conn,1,1
if not ( ssssrs.eof or ssssrs.bof ) then
%>
                <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                                      onmouseover="this.style.backgroundColor='#BECFDF';"> 
                <td height="18" align="center"><%=ssssrs("tname")%></td>
                <td align="center"><%=ssssrs("jersey")%></td>
                <td align="center"><%=ssssrs("pname")%></td>
                <td align="center"><%=amaxscorer(j-i)%></td>
              </tr>
<%
end if
ssssrs.close
set ssssrs=nothing
next
%>
            </table> 
<%
erase aplayer
erase amaxscorer
erase amaxpid
%>
<!-- ========================================================================================================		  

													hot shot  ============stop

 ======================================================================================================== -->
          </td>
        </tr>
      </table>
      <p>&nbsp;</p>
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
		  <td width="20">&nbsp;</td>
          <td> 
<!-- ========================================================================================================		  

													card ============star

 ======================================================================================================== -->
<%
psql="select distinct pid from card where tid=0 "
for i=0 to ubound(ateam)
psql=psql&" or tid="&ateam(i)
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
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' There are not player in The League(cup) of Worst Offenders ! ');</SCRIPT>")
response.End()
prs.close
set prs=nothing
end if
%>
<%
dim amaxred()
dim amaxyellow()
redim amaxred(ubound(aplayer))
redim amaxyellow(ubound(aplayer))
redim amaxpid(ubound(aplayer))
for i=0 to ubound(aplayer)


ssssql="select * from card where pid="&aplayer(i)
Set sssrs=Server.CreateObject("ADODB.RecordSet") 
sssrs.Open ssssql,conn,1,1
if not ( sssrs.eof or sssrs.bof ) then
for j=0 to sssrs.recordcount-1
amaxred(i)=amaxred(i)+clng(sssrs("red"))
amaxyellow(i)=amaxyellow(i)+clng(sssrs("yellow"))
sssrs.movenext
next
amaxpid(i)=aplayer(i)
end if
sssrs.close
set sssrs=nothing


next


j=ubound(amaxred)
for i=0 to ubound(amaxred)
for k=0 to j
if k=j then exit for
if amaxred(k)>amaxred(k+1) then
tmpmaxred=amaxred(k) 
amaxred(k)=amaxred(k+1)
amaxred(k+1)=tmpmaxred
tmpmaxyellow=amaxyellow(k) 
amaxyellow(k)=amaxyellow(k+1)
amaxyellow(k+1)=tmpmaxyellow
tmpmaxpid=amaxpid(k) 
amaxpid(k)=amaxpid(k+1)
amaxpid(k+1)=tmpmaxpid
end if
next
j=j-1
next
%>
            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#6184A3">
              <tr> 
                <td height="25" colspan="5" bgcolor="#FFFFFF"><b><%=cname%> Worst Offenders</b></td>
              </tr>
                <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                                      onmouseover="this.style.backgroundColor='#BECFDF';"> 
                <td width="190" height="20" align="center"><b>Team 
                  Name</b></td>
                <td width="45" align="center"><b>Jersey</b></td>
                <td align="center"><b>Name</b></td>
                <td width="45" align="center"><b>Yellow</b></td>
                <td width="45" align="center"><b>Red</b></td>
              </tr>
<%
if ubound(aplayer)>19 then 
j=19
else
j=ubound(aplayer)
end if
for i=0 to j
sssssql="select * from scorer where pid="&amaxpid(j-i)
Set ssssrs=Server.CreateObject("ADODB.RecordSet") 
ssssrs.Open sssssql,conn,1,1
if not ( ssssrs.eof or ssssrs.bof ) then
%>
                <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                                      onmouseover="this.style.backgroundColor='#BECFDF';"> 
                <td height="18" align="center"><%=ssssrs("tname")%></td>
                <td align="center"><%=ssssrs("jersey")%></td>
                <td align="center"><%=ssssrs("pname")%></td>
                <td align="center">
				<%
				if amaxyellow(j-i)=0 then
				 response.Write(amaxyellow(j-i))
				else
				%>
                  <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="70%" align="center"><%=amaxyellow(j-i)%> x</td>
                      <td width="30%"><img src="images/yellowcard.gif" width="10" height="18"></td>
                    </tr>
                  </table>
				<%
				end if
				%> 
				</td>
                <td align="center">
				<%
				if amaxred(j-i)=0 then
				 response.Write(amaxred(j-i))
				else
				%>
                  <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="70%" align="center"><%=amaxred(j-i)%> x</td>
                      <td width="30%"><img src="images/redcard.gif" width="10" height="18"></td>
                    </tr>
                  </table>
				<%
				end if
				%>  
				</td>
              </tr>
<%
end if
ssssrs.close
set ssssrs=nothing
next
%>
            </table> 
<%
erase aplayer
erase amaxred
erase amaxyellow
erase amaxpid
erase ateam
%>
<!-- ========================================================================================================		  

													card  ============stop

 ======================================================================================================== -->
          </td>
        </tr>
      </table>
      
    </td>
  </tr>
</table>
<!--#include file="Copyright.asp"-->