<SCRIPT LANGUAGE="JavaScript">
<!--

function getcidweek(targ,cid,selObj,restore){ //v3.0
  eval(targ+".location='league.asp?cid="+cid+"&mweek="+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}

//-->

</SCRIPT>

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
            <B></B> 
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr>
                <td width="158"><a href="League.asp"><b><font color="#1F4A71">CHOOSE A LEAGUE 
                  </font></b></a></td>
              </tr>
            </table> 
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <%
sql="select * from tournament where type='league' order by id_order" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
do while not rs.eof
%>
              <tr> 
                <td width="5" valign="top">
                  <img src="images/sun_l_sp.gif" width="5" height="10" border="0"> 
                  </td>
                <td><a href="League.asp?cid=<%=rs("cid")%>" title="<%=rs("name")%>" class="b"><%=rs("name")%></a></td>
              </tr>
              <%
rs.movenext
loop
rs.close
set rs=nothing
%>
            </table>
            <!-- ========================================================================================================		  

													cup table ============stop

 ======================================================================================================== -->
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="Cup.asp"><b><font color="#1F4A71">CUP MATCHES 
                  </font></b></a><font color="#1F4A71">&nbsp;+ +</font></td>
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
                <td><a href="TeamProfile.asp"><b><font color="#1F4A71">TEAMPORFILE</font></b></a> 
                  <font color="#1F4A71">&nbsp;+ +</font> </td>
              </tr>
            </table></td>
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
if cid="" then cid=210
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

<!-- --------------------------------------------------------------------------------------------------
													assign to "mweek"
---------------------------------------------------------------------------------------------------- -->
<%
msql="select mweek from match where cid="&cid&" order by mweek desc" 
Set mrs=Server.CreateObject("ADODB.RecordSet") 
mrs.Open msql,conn,1,1 
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#6184A3">
  <tr> 
    <td height="25" colspan="11" bgcolor="#FFFFFF"><b>
<%=cname%> -- Fixture and Result 
	</b></td>
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
  <tr> 
    <td colspan="11" bgcolor="#FFFFFF"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="29%"><b>Week : <%=mweek%></b></td>
          <td width="71%" align="right"><TABLE cellPadding=0 cellSpacing=0 width=100%>
              <TBODY>
                <TR> 
                              <TD width=234 align="right"><b><FONT color=#000066 
                        face="Verdana, Arial, Helvetica, sans-serif">Display Results 
                                For Week:</FONT></b></TD>
                  <TD width=114 align="right">
				  <select name="mweek" id="mweek" class="display_dropdown" onchange="getcidweek('self',<%=cid%>,this,0)">
				  <option selected>SELECT WEEK ..</option>
				  <option >--------------------------------------</option>

<%
dim amweek()
mwsql="select distinct mweek from match where cid="&cid&" order by mweek desc" 
Set mwrs=Server.CreateObject("ADODB.RecordSet") 
mwrs.Open mwsql,conn,1,1 
if not ( mwrs.eof or mwrs.bof ) then
redim amweek(mwrs.recordcount-1)
for i=0 to ubound(amweek)
amweek(i)=mwrs("mweek")
mwrs.movenext
next
mwrs.close
set mwrs=nothing
end if
%>
<%
if ubound(amweek)>0 then
for i=0 to ubound(amweek)
wsql="select * from match where mweek=" & amweek(i)
Set wrs=Server.CreateObject("ADODB.RecordSet")
wrs.Open wsql,conn,1,1
if not ( wrs.eof or wrs.bof ) then
maxday=wrs("mday")
maxmonth=wrs("mmonth")
maxyear=wrs("myear")
minday=wrs("mday")
minmonth=wrs("mmonth")
minyear=wrs("myear")
do while not wrs.eof
if wrs("mday")>maxday then maxday=wrs("mday")
if wrs("mmonth")>maxmonth then maxmonth=wrs("mmonth")
if wrs("myear")>maxyear then maxyear=wrs("myear")
if wrs("mday")<minday then minday=wrs("mday")
if wrs("mmonth")<minmonth then minmonth=wrs("mmonth")
if wrs("myear")<minyear then minyear=wrs("myear")
wrs.movenext
loop
end if
wrs.close
set wrs=nothing
%>
                      <option value=<%=amweek(i)%>>WEEK [<%=amweek(i)%>] , <%=minday%>/<%=minmonth%>/<%=minyear%>--<%=maxday%>/<%=maxmonth%>/<%=maxyear%></option>
<%
next
end if
erase amweek
%>
                  </select>
				  </TD>
                </TR>
              </TBODY>
            </TABLE></td>
        </tr>
      </table></td>
  </tr>
    <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                          onmouseover="this.style.backgroundColor='#BECFDF';"> 
    <td width="60" align="center"><b>Date</b></td>
    <td width="35" align="center"><b>Time</b></td>
    <td width="35" align="center"><b>Day</b></td>
    <td width="45" align="center"><b>Home Confirm</b></td>
    <td align="center"><b>Home Team</b></td>
    <td width="13" align="center"><b>CL</b></td>
    <td width="30" align="center"><b>VS</b></td>
    <td width="13"><b>CL</b></td>
    <td align="center"><b>Away Team</b></td>
    <td width="45" align="center"><b>Away Confirm</b></td>
    <td width="91" align="center"><b>Location</b></td>
  </tr>
<%
dim rrhscore(1)
dim rrascore(1)
sql="select * from match where mweek="&mweek&" and cid="&cid&" order by mweek desc ,mday desc" 
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.Open sql,conn,1,1
%>
<%
do while not rs.eof
matchdata=cstr(rs("mmonth"))&"/"&cstr(rs("mday"))&"/"&cstr(rs("myear"))
matchtime=cstr(rs("mhour"))&":"&cstr(rs("mminute"))
%>
  <tr bgcolor="FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
      onmouseover="this.style.backgroundColor='#BECFDF';"> 
    <td height="28" align="center"><%=matchdata%></td>
    <td align="center"><%=FormatDateTime(matchtime,4)%></td>
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
    <td align="center"><%=rs("htname")%>
	<br></td>
    <td bgcolor=<%=rs("hcolor")%>>&nbsp;</td>
    <td align="center">
	
<%
rrstr="no"
rrhscore(0)=100
rrascore(0)=100
rrhscore(1)=101
rrascore(1)=101

rsql="select * from rresult where mid="&rs("mid")
Set rrs=Server.CreateObject("ADODB.RecordSet")
rrs.Open rsql,conn,1,1
zi=0
do while not rrs.eof
rrhscore(zi)=rrs("hscore")
rrascore(zi)=rrs("ascore")
zi=zi+1
rrs.movenext
loop
rrs.close
set rrs=nothing
if isempty(rrhscore(0)) or isempty(rrhscore(1)) or isempty(rrascore(0)) or isempty(rrascore(1)) then
rrstr="no"
else
if rrhscore(0)=rrhscore(1) and rrascore(0)=rrascore(1) then rrstr="yes"
end if

if rrstr="yes" then
rresultstr=rrhscore(0)&":"&rrascore(0)
else
rresultstr="-"
end if
%>
<%
response.Write(rresultstr)
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
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
		  <td width="20">&nbsp;</td>
          <td>
<!-- ========================================================================================================		  

													standing ============star

 ======================================================================================================== -->
 <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#6184A3">
  <tr bgcolor="#FFFFFF"> 
    <td height="25" colspan="9" rowspan="2"><b><%=cname%> Standing 
      / (Tables)</b></td>
  </tr>
  <tr bgcolor="#FFFFFF"> </tr>
    <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                          onmouseover="this.style.backgroundColor='#BECFDF';"> 
    <td width="190" height="20" align="center"><b>Team Name</b></td>
    <td width="45" align="center"><b>P</b></td>
    <td width="45" align="center"><b>W</b></td>
    <td width="45" align="center"><b>D</b></td>
    <td width="45" align="center"><b>L</b></td>
    <td width="45" align="center"><b>F</b></td>
    <td width="45" align="center"><b>A</b></td>
    <td width="45" align="center"><b>GD</b></td>
    <td width="45" align="center"><b>PT</b></td>
  </tr>
<%
for i=0 to ubound(ateam)-1
sssql="select * from standing where tid="&ateam(i)
Set ssrs=Server.CreateObject("ADODB.RecordSet") 
ssrs.Open sssql,conn,1,1
if not ( ssrs.eof or ssrs.bof ) then
%>
  <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
      onmouseover="this.style.backgroundColor='#BECFDF';"> 
    <td height="18" align="center"><%=ssrs("tname")%></td>
    <td align="center"><%=ssrs("played")%></td>
    <td align="center"><%=ssrs("win")%></td>
    <td align="center"><%=ssrs("draw")%></td>
    <td align="center"><%=ssrs("lose")%></td>
    <td align="center"><%=ssrs("scored")%></td>
    <td align="center"><%=ssrs("conceded")%></td>
    <td align="center"><%=ssrs("gd")%></td>
    <td align="center"><%=ssrs("points")%></td>
  </tr>
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
      </table>
      <br>
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
		  <td width="20">&nbsp;</td>
          <td> 
<!-- ========================================================================================================		  

													hot shot  ============star

 ======================================================================================================== -->
<%
conn.execute "delete * from temscore"

psql="select pid,default_score,default_yellow,default_red from player where tid=0 "
for i=0 to ubound(ateam)
psql=psql&" or tid="&ateam(i)
next
ouinsertstr="insert into temscore (pid,score,yellow,red) " & psql
conn.Execute(ouinsertstr)
%>
<%
dim aplayer()
psql="select distinct pid from player where tid=0 "
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
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' There are not player in The League(cup) ! ');</SCRIPT>")
response.End()
prs.close
set prs=nothing
end if
%>
<%
Set sssrs=Server.CreateObject("ADODB.RecordSet")

for i=0 to ubound(aplayer)

ssssql="select * from scorer where pid="&aplayer(i)
sssrs.Open ssssql,conn,1,1
if not ( sssrs.eof or sssrs.bof ) then
for j=0 to sssrs.recordcount-1

rrstr="no"

rsql="select * from rresult where mid="&sssrs("mid")
Set rrs=Server.CreateObject("ADODB.RecordSet")
rrs.Open rsql,conn,1,1
zi=0
do while not rrs.eof
rrhscore(zi)=rrs("hscore")
rrascore(zi)=rrs("ascore")
zi=zi+1
rrs.movenext
loop
rrs.close
set rrs=nothing
if isempty(rrhscore(0)) or isempty(rrhscore(1)) or isempty(rrascore(0)) or isempty(rrascore(1)) then
rrstr="no"
else
if rrhscore(0)=rrhscore(1) and rrascore(0)=rrascore(1) then rrstr="yes"
end if

if rrstr="yes" then conn.Execute("insert into temscore (pid,score) select pid,score from scorer where pid="&aplayer(i))

sssrs.movenext
next

end if
sssrs.close
next












for i=0 to ubound(aplayer)

ssssql="select * from card where pid="&aplayer(i)
sssrs.Open ssssql,conn,1,1
if not ( sssrs.eof or sssrs.bof ) then
for j=0 to sssrs.recordcount-1

rrstr="no"

rsql="select * from rresult where mid="&sssrs("mid")
Set rrs=Server.CreateObject("ADODB.RecordSet")
rrs.Open rsql,conn,1,1
zi=0
do while not rrs.eof
rrhscore(zi)=rrs("hscore")
rrascore(zi)=rrs("ascore")
zi=zi+1
rrs.movenext
loop
rrs.close
set rrs=nothing
if isempty(rrhscore(0)) or isempty(rrhscore(1)) or isempty(rrascore(0)) or isempty(rrascore(1)) then
rrstr="no"
else
if rrhscore(0)=rrhscore(1) and rrascore(0)=rrascore(1) then rrstr="yes"
end if

if rrstr="yes" then conn.Execute("insert into temscore (pid,yellow,red) select pid,yellow,red from card where pid="&aplayer(i))

sssrs.movenext
next

end if
sssrs.close
next


set sssrs=nothing

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
sql="SELECT pid,sum(score) as sumscore FROM temscore group by pid order by sum(score) desc;"
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,3
i=0
do while not rs.eof

psql="select name,jersey,tid from player where pid=" & rs("pid")
set prs=server.createobject("ADODB.Recordset")
prs.open psql,conn,1,3
if not prs.eof then

if not isnull(prs("tid")) then 
tsql="select name from team where tid=" & prs("tid")
set trs=server.createobject("ADODB.Recordset")
trs.open tsql,conn,1,3
if not trs.eof then
tname=trs("name")
end if
trs.close
set trs=nothing
end if

pname=prs("name")
jersey=prs("jersey")
end if
prs.close
set prs=nothing


%>
           <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                                 onmouseover="this.style.backgroundColor='#BECFDF';"> 
                <td height="18" align="center"><%=tname%>&nbsp;</td>
                <td align="center"><%=jersey%>&nbsp;</td>
                <td align="center"><%=pname%>&nbsp;</td>
                <td align="center"><%=rs("sumscore")%></td>
              </tr>
<%
i=i+1
if i>10 then exit do
rs.movenext
loop
rs.close
%>
            </table> 
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
            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#6184A3">
              <tr bgcolor="#FFFFFF"> 
                <td height="25" colspan="5"><b><%=cname%> Worst Offenders</b></td>
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
sql="SELECT pid,sum(red) as sumred,sum(yellow) as sumyellow FROM temscore group by pid order by sum(red) desc ,sum(yellow) desc;"
rs.Open sql,conn,1,1
i=0
do while not rs.eof



psql="select name,jersey,tid from player where pid=" & rs("pid")
set prs=server.createobject("ADODB.Recordset")
prs.open psql,conn,1,1
if not prs.eof then

if not isnull(prs("tid")) then 
tsql="select name from team where tid=" & prs("tid")
set trs=server.createobject("ADODB.Recordset")
trs.open tsql,conn,1,1
if not trs.eof then
tname=trs("name")
end if
trs.close
set trs=nothing
end if

pname=prs("name")
jersey=prs("jersey")
end if
prs.close
set prs=nothing


%>
           <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                                 onmouseover="this.style.backgroundColor='#BECFDF';"> 
                <td height="18" align="center"><%=tname%></td>
                <td align="center"><%=jersey%></td>
                <td align="center"><%=pname%></td>
                <td align="center">
				<%
				if rs("sumyellow")=0 then
				 response.Write(rs("sumyellow"))
				else
				%>
                  <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="70%" align="center"><%=rs("sumyellow")%> x</td>
                      <td width="30%"><img src="images/yellowcard.gif" width="10" height="18"></td>
                    </tr>
                  </table>
				<%
				end if
				%>
				</td>
                <td align="center"> 
				<%
				if rs("sumred")=0 then
				 response.Write(rs("sumred"))
				else
				%>
                  <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="70%" align="center"><%=rs("sumred")%> x</td>
                      <td width="30%"><img src="images/redcard.gif" width="10" height="18"></td>
                    </tr>
                  </table>
				<%
				end if
				%>  
				  </td>
              </tr>
<%
i=i+1
if i>10 then exit do
rs.movenext
loop
rs.close
set rs=nothing
%>
            </table> 
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