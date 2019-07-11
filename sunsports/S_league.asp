<html>
<head>
<title>sunsports.com.sg</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<LINK href="Css.css" type=text/css rel=stylesheet>
<LINK href="favicon.ico" rel="SHORTCUT ICON">
</head>

<body leftmargin="0" topmargin="0">
<!--#include file="Top.asp"-->
<table width="776" height="100%" border="1" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="181" valign="top"> <br>
      <table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td>
            <!--#include file="League_Table.asp" -->
          </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table>
    </td>
    <td width="579" valign="top"><br>
      <table width="96%" border="0" align="center" cellpadding="2" cellspacing="2">
        <tr>
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
redim ateam(trs.recordcount)
for i=0 to ubound(ateam)-1
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
<table width="96%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FF6633">
  <tr bgcolor="#FFFFFF"> 
    <td height="25" colspan="11"><b>
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
  <tr bgcolor="#FFFFFF"> 
    <td colspan="11"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="34%"><b>Week : <%=mweek%></b></td>
          <td width="66%" align="right"><TABLE cellPadding=0 cellSpacing=0 width=350>
              <TBODY>
                <TR> 
                  <TD width=234 align="right"><b><FONT color=#000066 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        size=2>Display Results For Week:</FONT></b></TD>
                  <TD width=114 align="right">
				  <select name=weeks class="display_dropdown">
<%
dim amweek()
mwsql="select distinct mweek from match order by mweek desc" 
Set mwrs=Server.CreateObject("ADODB.RecordSet") 
mwrs.Open mwsql,conn,1,1 
if not ( mwrs.eof or mwrs.bof ) then
redim amweek(mwrs.recordcount-1)
for i=0 to ubound(amweek)-1
amweek(i)=mwrs("mweek")
mwrs.movenext
next
mwrs.close
set mwrs=nothing
end if
%>
<%
if ubound(amweek)>0 then
for i=0 to ubound(amweek)-1
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
                      <option selected value=<%=amweek(i)%>>==<%=amweek(i)%>==<%=minday%>/<%=minmonth%>/<%=minyear%>--<%=maxday%>/<%=maxmonth%>/<%=maxyear%></option>
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
  <tr> 
    <td width="56" align="center" bgcolor="#FFFFFF"><b>Date</b></td>
    <td width="56" align="center" bgcolor="#FFFFFF"><b>Time</b></td>
    <td width="38" align="center" bgcolor="#FFFFFF"><b>Day</b></td>
    <td width="89" align="center" bgcolor="#FFFFFF"><b>Home Confirm</b></td>
    <td width="124" align="center" bgcolor="#FFFFFF"><b>Home Team</b></td>
    <td width="25" align="center" bgcolor="#FFFFFF"><b>CL</b></td>
    <td width="71" align="center" bgcolor="#FFFFFF"><b>VS</b></td>
    <td width="25" bgcolor="#FFFFFF"><b>CL</b></td>
    <td width="147" align="center" bgcolor="#FFFFFF"><b>Away Team</b></td>
    <td width="89" align="center" bgcolor="#FFFFFF"><b>Away Confirm</b></td>
    <td width="224" align="center" bgcolor="#FFFFFF"><b>Location</b></td>
  </tr>
<%
sql="select * from match where mweek="&mweek&" and cid="&cid&" order by myear,mmonth,mday,mhour,mminute desc" 
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.Open sql,conn,1,1
%>
<%
do while not rs.eof
matchdata=cstr(rs("mmonth"))&"-"&cstr(rs("mday"))&"-"&cstr(rs("myear"))
matchtime=cstr(rs("mhour"))&":"&cstr(rs("mminute"))
%>
  <tr> 
    <td height="35" align="center" bgcolor="#f5f5f5"><%=matchdata%> </td>
    <td align="center" bgcolor="#f5f5f5"><%=matchtime%></td>
    <td align="center" bgcolor="#f5f5f5"><%=weekdayname(weekday(matchdata))%></td>
    <td align="center" bgcolor="#f5f5f5"><%=rs("hconfirm")%></td>
    <td height="35" align="center" bgcolor="#f5f5f5"><%=rs("htname")%>
	<br></td>
    <td bgcolor=<%=rs("hcolor")%>>&nbsp;</td>
    <td align="center" bgcolor="#FFFFFF">
	
<% if trim(rs("matched"))="yes" then %>
<!-- --------------------------------------------------------------------------------------------------
													assign to "hsumscorer or asumscorer"
---------------------------------------------------------------------------------------------------- -->
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
    <td align="center" bgcolor="#f5f5f5"><%=rs("atname")%></td>
    <td align="center" bgcolor="#f5f5f5"><%=rs("aconfirm")%></td>
    <td align="center" bgcolor="#f5f5f5"><%=rs("fname")%></td>
  </tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
  <tr> 
    <td height="35" align="center" bgcolor="#f5f5f5">07/03</td>
    <td bgcolor="#f5f5f5">15:30</td>
    <td align="center" bgcolor="#f5f5f5">Sun</td>
    <td align="center" bgcolor="#f5f5f5">Yes</td>
    <td height="35" align="center" bgcolor="#f5f5f5">Barbarians </td>
    <td bgcolor="#FFFF00">&nbsp;</td>
    <td align="center" bgcolor="#FFFFFF">-</td>
    <td bgcolor="#0099FF">&nbsp;</td>
    <td align="center" bgcolor="#f5f5f5">Hibernians A </td>
    <td align="center" bgcolor="#f5f5f5">No</td>
    <td bgcolor="#f5f5f5">Sentosa West Field </td>
  </tr>
</table>
		  
<!-- ========================================================================================================		  

													fixture or result ============stop

 ======================================================================================================== -->		  
		  
          </td>
        </tr>
      </table> <br>
      <table width="96%" border="0" align="center" cellpadding="2" cellspacing="2">
        <tr> 
          <td>
<!-- ========================================================================================================		  

													standing ============star

 ======================================================================================================== -->
 <table width="96%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#660000">
  <tr bgcolor="#FFFFFF"> 
    <td height="25" colspan="9" rowspan="2"><b><%=cname%> Standing 
      / (Tables)</b></td>
  </tr>
  <tr bgcolor="#FFFFFF"> </tr>
  <tr> 
    <td width="190" height="20" align="center" bgcolor="#FFFFFF"><b>Team Name</b></td>
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
      <br>
      <table width="96%" border="0" align="center" cellpadding="2" cellspacing="2">
        <tr> 
          <td> 
<!-- ========================================================================================================		  

													hot shot  ============star

 ======================================================================================================== -->
<%
dim aplayer()
psql="select distinct pid from scorer where tid=0 "
for i=0 to ubound(ateam)-1
psql=psql&" or tid="&ateam(i)
next
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1 
if not ( prs.eof or prs.bof ) then
redim aplayer(prs.recordcount)
for i=0 to ubound(aplayer)-1
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
<%
for i=0 to ubound(aplayer)-1
ssssql="select * from scorer where pid="&aplayer(i)
Set sssrs=Server.CreateObject("ADODB.RecordSet") 
sssrs.Open ssssql,conn,1,1
if not ( sssrs.eof or sssrs.bof ) then
do 
maxscorer=maxscorer+clng(sssrs("score"))
sssrs.movenext
loop while sssrs.eof 
%>
              <tr> 
                <td align="center" bgcolor="#FFFFFF"><%=sssrs("tname")%></td>
                <td align="center" bgcolor="#FFFFFF"><%=sssrs("jersey")%></td>
                <td align="center" bgcolor="#FFFFFF"><%=sssrs("pname")%></td>
                <td align="center" bgcolor="#FFFFFF"><%=maxscorer%></td>
              </tr>
<%
end if
sssrs.close
set sssrs=nothing
next
%>
              <tr> 
                <td align="center" bgcolor="#FFFFFF">
				BNP Paribas</td>
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
      <p>&nbsp;</p>
      <table width="96%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#660000">
        <tr bgcolor="#FFFFFF"> 
          <td height="25" colspan="5"><b>Sunsports Champions League's Worst Offenders</b></td>
        </tr>
        <tr> 
          <td width="190" height="20" align="center" bgcolor="#FFFFFF"><b>Team Name</b></td>
          <td width="45" align="center" bgcolor="#FFFFFF"><b>Jersey</b></td>
          <td align="center" bgcolor="#FFFFFF"><b>Name</b></td>
          <td width="45" align="center" bgcolor="#FFCC00"><b>Yellow</b></td>
          <td width="45" align="center" bgcolor="#FF0000"><b>Red</b></td>
        </tr>
        <tr> 
          <td align="center" bgcolor="#FFFFFF">BNP Paribas</td>
          <td align="center" bgcolor="#FFFFFF">25</td>
          <td align="center" bgcolor="#FFFFFF">Khairul Sufian</td>
          <td align="center" bgcolor="#FFFFFF">2</td>
          <td align="center" bgcolor="#FFFFFF">1</td>
        </tr>
        <tr> 
          <td align="center" bgcolor="#FFFFFF">Racing Carnegies FC</td>
          <td align="center" bgcolor="#FFFFFF">24</td>
          <td align="center" bgcolor="#FFFFFF">Jumain Bin Semin</td>
          <td align="center" bgcolor="#FFFFFF">2</td>
          <td align="center" bgcolor="#FFFFFF">1</td>
        </tr>
        <tr> 
          <td align="center" bgcolor="#FFFFFF">Racing Carnegies FC</td>
          <td align="center" bgcolor="#FFFFFF">24</td>
          <td align="center" bgcolor="#FFFFFF">Ronald Chan</td>
          <td align="center" bgcolor="#FFFFFF">2</td>
          <td align="center" bgcolor="#FFFFFF">1</td>
        </tr>
        <tr> 
          <td align="center" bgcolor="#FFFFFF">BNP Paribas</td>
          <td align="center" bgcolor="#FFFFFF">24</td>
          <td align="center" bgcolor="#FFFFFF">Mat Boylan</td>
          <td align="center" bgcolor="#FFFFFF">2</td>
          <td align="center" bgcolor="#FFFFFF">1</td>
        </tr>
        <tr> 
          <td align="center" bgcolor="#FFFFFF">BNP Paribas</td>
          <td align="center" bgcolor="#FFFFFF">24</td>
          <td align="center" bgcolor="#FFFFFF">Mark Fyye</td>
          <td align="center" bgcolor="#FFFFFF">2</td>
          <td align="center" bgcolor="#FFFFFF">1</td>
        </tr>
        <tr> 
          <td align="center" bgcolor="#FFFFFF">BNP Paribas</td>
          <td align="center" bgcolor="#FFFFFF">24</td>
          <td align="center" bgcolor="#FFFFFF">David Benntte</td>
          <td align="center" bgcolor="#FFFFFF">2</td>
          <td align="center" bgcolor="#FFFFFF">1</td>
        </tr>
        <tr> 
          <td align="center" bgcolor="#FFFFFF">BNP Paribas</td>
          <td align="center" bgcolor="#FFFFFF">24</td>
          <td align="center" bgcolor="#FFFFFF">Peter Yong</td>
          <td align="center" bgcolor="#FFFFFF">2</td>
          <td align="center" bgcolor="#FFFFFF">1</td>
        </tr>
        <tr> 
          <td align="center" bgcolor="#FFFFFF">BNP Paribas</td>
          <td align="center" bgcolor="#FFFFFF">24</td>
          <td align="center" bgcolor="#FFFFFF">Zanil</td>
          <td align="center" bgcolor="#FFFFFF">2</td>
          <td align="center" bgcolor="#FFFFFF">1</td>
        </tr>
        <tr> 
          <td align="center" bgcolor="#FFFFFF">BNP Paribas</td>
          <td align="center" bgcolor="#FFFFFF">24</td>
          <td align="center" bgcolor="#FFFFFF">Richard</td>
          <td align="center" bgcolor="#FFFFFF">2</td>
          <td align="center" bgcolor="#FFFFFF">1</td>
        </tr>
        <tr> 
          <td align="center" bgcolor="#FFFFFF">BNP Paribas</td>
          <td align="center" bgcolor="#FFFFFF">24</td>
          <td align="center" bgcolor="#FFFFFF">Mohd Afdzal Bin Mohd Ibrahim</td>
          <td align="center" bgcolor="#FFFFFF">2</td>
          <td align="center" bgcolor="#FFFFFF">1</td>
        </tr>
        <tr> 
          <td align="center" bgcolor="#FFFFFF">BNP Paribas</td>
          <td align="center" bgcolor="#FFFFFF">24</td>
          <td align="center" bgcolor="#FFFFFF">Lee Taylor</td>
          <td align="center" bgcolor="#FFFFFF">2</td>
          <td align="center" bgcolor="#FFFFFF">1</td>
        </tr>
        <tr> 
          <td align="center" bgcolor="#FFFFFF">BNP Paribas</td>
          <td align="center" bgcolor="#FFFFFF">24</td>
          <td align="center" bgcolor="#FFFFFF">Stephen Begley</td>
          <td align="center" bgcolor="#FFFFFF">2</td>
          <td align="center" bgcolor="#FFFFFF">1</td>
        </tr>
        <tr> 
          <td align="center" bgcolor="#FFFFFF">BNP Paribas</td>
          <td align="center" bgcolor="#FFFFFF">24</td>
          <td align="center" bgcolor="#FFFFFF">Ronald</td>
          <td align="center" bgcolor="#FFFFFF">2</td>
          <td align="center" bgcolor="#FFFFFF">1</td>
        </tr>
        <tr> 
          <td align="center" bgcolor="#FFFFFF">BNP Paribas</td>
          <td align="center" bgcolor="#FFFFFF">24</td>
          <td align="center" bgcolor="#FFFFFF">Lee Taylor</td>
          <td align="center" bgcolor="#FFFFFF">2</td>
          <td align="center" bgcolor="#FFFFFF">1</td>
        </tr>
        <tr> 
          <td align="center" bgcolor="#FFFFFF">BNP Paribas</td>
          <td align="center" bgcolor="#FFFFFF">24</td>
          <td align="center" bgcolor="#FFFFFF">Lee Taylor</td>
          <td align="center" bgcolor="#FFFFFF">2</td>
          <td align="center" bgcolor="#FFFFFF">1</td>
        </tr>
        <tr> 
          <td align="center" bgcolor="#FFFFFF">BNP Paribas</td>
          <td align="center" bgcolor="#FFFFFF">24</td>
          <td align="center" bgcolor="#FFFFFF">Lee Taylor</td>
          <td align="center" bgcolor="#FFFFFF">2</td>
          <td align="center" bgcolor="#FFFFFF">1</td>
        </tr>
        <tr> 
          <td align="center" bgcolor="#FFFFFF">BNP Paribas</td>
          <td align="center" bgcolor="#FFFFFF">24</td>
          <td align="center" bgcolor="#FFFFFF">Lee Taylor</td>
          <td align="center" bgcolor="#FFFFFF">2</td>
          <td align="center" bgcolor="#FFFFFF">1</td>
        </tr>
        <tr> 
          <td align="center" bgcolor="#FFFFFF">BNP Paribas</td>
          <td align="center" bgcolor="#FFFFFF">24</td>
          <td align="center" bgcolor="#FFFFFF">Lee Taylor</td>
          <td align="center" bgcolor="#FFFFFF">2</td>
          <td align="center" bgcolor="#FFFFFF">1</td>
        </tr>
        <tr> 
          <td align="center" bgcolor="#FFFFFF">BNP Paribas</td>
          <td align="center" bgcolor="#FFFFFF">24</td>
          <td align="center" bgcolor="#FFFFFF">Lee Taylor</td>
          <td align="center" bgcolor="#FFFFFF">2</td>
          <td align="center" bgcolor="#FFFFFF">1</td>
        </tr>
      </table>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <p>&nbsp;</p></td>
  </tr>
</table>
<!--#include file="Copyright.asp"-->
</body>
</html>
