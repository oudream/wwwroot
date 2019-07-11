<!--#include file="conn/conn.asp" -->
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
if mweek="" then mweek=rs("mweek")
mrs.close
set mrs=nothing
%>
  <tr bgcolor="#FFFFFF"> 
    <td colspan="11"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="34%"><b>Week <%=mweek%> (07/02 - 08/02) </b></td>
          <td width="66%" align="right"><TABLE cellPadding=0 cellSpacing=0 width=350>
              <TBODY>
                <TR> 
                  <TD width=234 align="right"><b><FONT color=#000066 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        size=2>Display Results For Week:</FONT></b></TD>
                  <TD width=114 align="right">
				  <select name=weeks class="display_dropdown">
<%
mwsql="select distinct mweek from match order by mweek desc" 
Set mwrs=Server.CreateObject("ADODB.RecordSet") 
mwrs.Open mwsql,conn,1,1 
do while not mwrs.eof
wsql="select * from match where mweek="&mwrs("mweek")&" order by mweek desc" 
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
                      <option selected value=<%=mwrs("mweek")%>><%=minday%>/<%=minmonth%>/<%=minyear%>--<%=maxday%>/<%=maxmonth%>/<%=maxyear%></option>
<%
mwrs.movenext
loop
mwrs.close
set mwrs=nothing
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
<%
tsql="select * from team order by tid desc"
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1
if trs.eof or trs.bof then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Database is empty  ! ');</SCRIPT>")
response.End()
trs.close
set trs=nothing
end if
do while not trs.eof
if rs("htid")=trs("tid") then htname=trs("name")
if rs("atid")=trs("tid") then atname=trs("name")
trs.movenext
loop
trs.close
set trs=nothing
%>
  <tr> 
    <td height="35" align="center" bgcolor="#f5f5f5"><%=matchdata%> </td>
    <td align="center" bgcolor="#f5f5f5"><%=matchtime%></td>
    <td align="center" bgcolor="#f5f5f5"><%=weekdayname(weekday(matchdata))%></td>
    <td align="center" bgcolor="#f5f5f5"><%=rs("hconfirm")%></td>
    <td height="35" align="center" bgcolor="#f5f5f5"><%=htname%>
	<br></td>
    <td bgcolor=<%=rs("hcolor")%>>&nbsp;</td>
    <td align="center" bgcolor="#FFFFFF"><%=rs("matchdata")%></td>
    <td bgcolor=<%=rs("acolor")%>>&nbsp;</td>
    <td align="center" bgcolor="#f5f5f5"><%=atname%></td>
    <td align="center" bgcolor="#f5f5f5"><%=rs("matchdata")%></td>
    <td align="center" bgcolor="#f5f5f5"><%=rs("matchdata")%></td>
  </tr>
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
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
