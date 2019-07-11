<!--#include file="Top_sun.asp"-->
<table width="776" border="0" align="center" cellpadding="0" cellspacing="0">
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
            <B></B> <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td width="158"><a href="Fixtures-Results.asp"><b><font color="#1F4A71">CHOOSE 
                  A LEAGUE </font></b></a></td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <%
sql="select * from tournament where type='league'" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
do while not rs.eof
%>
              <tr> 
                <td width="5" valign="top"> <img src="images/sun_l_sp.gif" width="5" height="10" border="0"> 
                </td>
                <td><a href="Fixtures-Results.asp?cid=<%=rs("cid")%>" title="<%=rs("name")%>" class="b"><%=rs("name")%></a></td>
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
                <td><a href="Cup.asp"><b><font color="#1F4A71">CUP MATCHES </font></b></a><font color="#1F4A71">&nbsp;+ 
                  +</font></td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="Friendlies.asp"><b><font color="#1F4A71">FRIENDLY 
                  MATCHES </font></b></a></td>
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
      </table></td>
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
dim rrhscore(1)
dim rrascore(1)
sql="select top 20 * from match where cid="&cid&" order by mid desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#6184A3">
  <tr> 
    <td height="25" colspan="11" bgcolor="#FFFFFF"><b>
<%=cname%> -- Recent Fixture and Result 
	</b></td>
  </tr>
<%
if rs.eof or rs.bof then
response.Write("<tr bgcolor=#FFFFFF><td height=25 colspan=11><b>NO DATA</b></td></tr></table>")
response.End()
rs.close
set rs=nothing
else
%>
    <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                          onmouseover="this.style.backgroundColor='#BECFDF';"> 
    <td width="60" align="center"><b>Date</b></td>
    <td width="35" align="center"><b>Time</b></td>
    <td width="35" align="center"><b>Day</b></td>
    <td width="45" align="center"><b>Home Confirm</b></td>
    <td align="center"><b>Home Team</b></td>
    <td width="13" align="center"><b>CL</b></td>
    <td width="71" align="center"><b>VS</b></td>
    <td width="13"><b>CL</b></td>
    <td align="center"><b>Away Team</b></td>
    <td width="45" align="center"><b>Away Confirm</b></td>
    <td width="91" align="center"><b>Location</b></td>
  </tr>
<%
do while not rs.eof
matchdata=cstr(rs("mmonth"))&"/"&cstr(rs("mday"))&"/"&cstr(rs("myear"))
matchtime=cstr(rs("mhour"))&":"&cstr(rs("mminute"))
%>
    <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                          onmouseover="this.style.backgroundColor='#BECFDF';"> 
    <td height="25" align="center"><%=matchdata%> </td>
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
    <td align="center"><%=rs("htname")%>
   </td>
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
end if
%>
</table>
		  
<!-- ========================================================================================================		  

													fixture or result ============stop

 ======================================================================================================== -->		  
		  
          </td>
        </tr>
      </table> 
    </td>
  </tr>
</table>
<!--#include file="Copyright.asp"-->