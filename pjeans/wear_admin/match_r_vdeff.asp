<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file="adminconn.asp" --> 
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr>
    <td>&nbsp;</td>
    <td height="50" align="center"><H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066><br>
        Match Report</FONT></H2></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><table width="100%" border="0" align="center" cellpadding="2" cellspacing="2">
        <tr> 
          <td> 
<%
sql="select * from match where cid<>0 order by mweek desc , mday desc" 
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.Open sql,conn,1,1
%>
            <table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
              <tr bgcolor="#FFFFFF"> 
                <td height="50" colspan="13"><b> Manager have not refered report</b></td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td colspan="13">&nbsp; </td>
              </tr>
              <tr> 
                <td width="49" align="center" bgcolor="#FFFFFF"><b>Week</b></td>
                <td width="49" align="center" bgcolor="#FFFFFF"><b>Date</b></td>
                <td width="49" align="center" bgcolor="#FFFFFF"><b>Time</b></td>
                <td width="34" align="center" bgcolor="#FFFFFF"><b>Day</b></td>
                <td width="168" align="center" bgcolor="#FFFFFF"><b>Home Team</b></td>
                <td width="19" align="center" bgcolor="#FFFFFF"><b>CL</b></td>
                <td width="18" align="center" bgcolor="#FFFFFF"><b>VS</b></td>
                <td width="18" bgcolor="#FFFFFF"><b>CL</b></td>
                <td width="172" align="center" bgcolor="#FFFFFF"><b>Away Team</b></td>
                <td width="120" align="center" bgcolor="#FFFFFF"><b>Location</b></td>
                <td width="120" align="center" bgcolor="#FFFFFF"><b>Operation</b></td>
              </tr>
              <%
if rs.eof then
response.Write("<tr bgcolor=#FFFFFF><td height=25 colspan=11><b> Managers have refered all report </b></td></tr></table>")
rs.close
set rs=nothing
response.End()
end if
%>
              <%
dim rrhscore(1)
dim rrascore(1)
ouifdata="no"

do while not rs.eof

hadded="no"
aadded="no"
rrstr="no"
rrhscore(0)=100
rrascore(0)=100
rrhscore(1)=101
rrascore(1)=101

rsql="select * from rresult where mid="&rs("mid")
Set rrs=Server.CreateObject("ADODB.RecordSet")
rrs.Open rsql,conn,1,1
do while not rrs.eof
if rrs("tid")=rs("htid") then
hadded="yes"
rrhscore(0)=rrs("hscore")
rrascore(0)=rrs("ascore")
end if
if rrs("tid")=rs("atid") then 
aadded="yes"
rrhscore(1)=rrs("hscore")
rrascore(1)=rrs("ascore")
end if
rrs.movenext
loop
rrs.close
set rrs=nothing
if rrhscore(0)=rrhscore(1) and rrascore(0)=rrascore(1) then rrstr="yes"

if rrstr="no" and hadded="yes" and aadded="yes" then 

ouifdata="yes"
matchdata=cstr(rs("mmonth"))&"/"&cstr(rs("mday"))&"/"&cstr(rs("myear"))
matchtime=cstr(rs("mhour"))&":"&cstr(rs("mminute"))
%>
              <tr> 
                <td height="35" align="center" bgcolor="#f5f5f5"><%=rs("mweek")%>&nbsp;</td>
                <td height="35" align="center" bgcolor="#f5f5f5"><%=matchdata%>&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5"><%=matchtime%>&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5"> <%
if weekday(matchdata,1)=1 then response.Write("SUN")
if weekday(matchdata,1)=2 then response.Write("MON")
if weekday(matchdata,1)=3 then response.Write("TUE")
if weekday(matchdata,1)=4 then response.Write("WED")
if weekday(matchdata,1)=5 then response.Write("THU")
if weekday(matchdata,1)=6 then response.Write("FRI")
if weekday(matchdata,1)=7 then response.Write("SAT")
%> </td>
                <td height="35" align="center" bgcolor="#f5f5f5"> <%=rs("htname")%><br> <font color="#FF0000"><%=rrhscore(0)&" : "&rrascore(0)%></font>&nbsp;</td>
                <td bgcolor=<%=rs("hcolor")%>>&nbsp;</td>
                <td align="center" bgcolor="#FFFFFF">-</td>
                <td bgcolor=<%=rs("acolor")%>>&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5"> <%=rs("atname")%><br><font color="#FF0000">
                  <font color="#FF0000"><%=rrhscore(1)&" : "&rrascore(1)%> </font>&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5"><%=rs("fname")%>&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5"><strong><a href="match_r_edeff.asp?mid=<%=rs("mid")%>">EDIT</a></strong>&nbsp;</td>
              </tr>
              <%
end if
rs.movenext
loop
rs.close
set rs=nothing
%>
<%
if ouifdata="no" then
%>
              <tr> 
                <td height="35" colspan="13" align="center" bgcolor="#f5f5f5"><strong>ALL 
                  REPORT IS SAME</strong></td>
              </tr>
<%
end if
%>
              <tr bgcolor="#FFFFFF">
                <td height="50" colspan="13" align="center">&nbsp;</td>
              </tr>
            </table>
            <!-- ========================================================================================================		  

													fixture or result ============stop

 ======================================================================================================== -->
          </td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
</table>
