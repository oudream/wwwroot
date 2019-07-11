<SCRIPT LANGUAGE="JavaScript">
<!--

function getweek(targ,selObj,restore){ //v3.0
  eval(targ+".location='match_r_no.asp?mweek="+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}

//-->

</SCRIPT>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#include file="adminconn.asp" --> 
<%
if request("option")="creatematchissucc" then response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Create match is complete. ');</SCRIPT>")
if request("option")="edit" then
if request("ha")="h" then sql="update match set hconfirm='"&request("tfstr")&"' where mid=" & request("mid")
if request("ha")="a" then sql="update match set aconfirm='"&request("tfstr")&"' where mid=" & request("mid")
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Update Match is complete. ');</SCRIPT>")
end if
%>
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr>
    <td>&nbsp;</td>
    <td><h2><font color="#000066"><br>
        Match Report<br>
        </font></h2></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><table width="100%" border="0" align="center" cellpadding="2" cellspacing="2">
        <tr> 
          <td> 
            <!-- ========================================================================================================		  

													fixture or result ============star

 ======================================================================================================== -->
            <!-- --------------------------------------------------------------------------------------------------
													assign to "cid"
---------------------------------------------------------------------------------------------------- -->
            <!-- --------------------------------------------------------------------------------------------------
													assign to "mweek"
---------------------------------------------------------------------------------------------------- -->
<%
msql="select mweek from match where htid="&session("manager_tid")&" or atid="&session("manager_tid")&" order by mday desc" 
Set mrs=Server.CreateObject("ADODB.RecordSet") 
mrs.Open msql,conn,1,1 
%> 
            <table width="100%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolor="#CCCCCC">
              <tr bgcolor="#FFFFFF"> 
                <td height="25" colspan="14"><b><font color="#FF0000"><%=session("manager_tname")%></font> match report</b></td>
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
                <td colspan="14">&nbsp; </td>
              </tr>
              <tr> 
                <td width="120" align="center" bgcolor="#FFFFFF"><strong>Tournament</strong></td>
                <td width="30" align="center" bgcolor="#FFFFFF"><strong>Week</strong></td>
                <td width="80" align="center" bgcolor="#FFFFFF"><b>Date</b></td>
                <td width="30" align="center" bgcolor="#FFFFFF"><b>Time</b></td>
                <td width="30" align="center" bgcolor="#FFFFFF"><b>Day</b></td>
                <td width="46" align="center" bgcolor="#FFFFFF"><b>Home Confirm</b></td>
                <td width="42" align="center" bgcolor="#FFFFFF"><b>Home Team</b></td>
                <td width="17" align="center" bgcolor="#FFFFFF"><b>CL</b></td>
                <td width="22" align="center" bgcolor="#FFFFFF"><b>VS</b></td>
                <td width="17" bgcolor="#FFFFFF"><b>CL</b></td>
                <td width="45" align="center" bgcolor="#FFFFFF"><b>Away Team</b></td>
                <td width="53" align="center" bgcolor="#FFFFFF"><b>Away Confirm</b></td>
                <td width="71" align="center" bgcolor="#FFFFFF"><b>Location</b></td>
                <td width="94" align="center" bgcolor="#FFFFFF"><strong>Operation</strong></td>
              </tr>
              
              <%
sql="select * from match where cid<>0 and( htid="&session("manager_tid")&" or atid="&session("manager_tid")&") order by mweek desc ,mday desc" 
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.Open sql,conn,1,1
if rs.eof or rs.bof then 
response.Write("<tr><tdcolspan=14 align=center bgcolor=#FFFFFF><strong>There are not match</strong></td></tr>")
response.End()
end if
%>
<%
ouifdata="no"
do while not rs.eof
if rs("htid")=session("manager_tid") then ha="h"
if rs("atid")=session("manager_tid") then ha="a"
matchdata=cstr(rs("mmonth"))&"/"&cstr(rs("mday"))&"/"&cstr(rs("myear"))
matchtime=cstr(rs("mhour"))&":"&cstr(rs("mminute"))
if DateDiff("d", date, cdate(matchdata))<=0 then
%>
              <%
ssql="select * from rresult where mid="&rs("mid")&" and tid="&session("manager_tid")
Set srs=Server.CreateObject("ADODB.RecordSet") 
srs.Open ssql,conn,1,1
if not ( srs.eof or srs.bof ) then
tfedit="yes"
srs.close
set srs=nothing
else
tfedit="no"
srs.close
set srs=nothing
end if
if tfedit="no" then 
ouifdata="yes"
%>
              <tr> 
                <td align="center" bgcolor="#f5f5f5"><%=rs("cname")%>&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5"><%=rs("mweek")%>&nbsp;</td>
                <td height="35" align="center" bgcolor="#f5f5f5"><%=matchdata%> </td>
                <td align="center" bgcolor="#f5f5f5"><%=matchtime%></td>
                <td align="center" bgcolor="#f5f5f5"> <%
if weekday(matchdata,1)=1 then response.Write("SUN")
if weekday(matchdata,1)=2 then response.Write("MON")
if weekday(matchdata,1)=3 then response.Write("TUE")
if weekday(matchdata,1)=4 then response.Write("WED")
if weekday(matchdata,1)=5 then response.Write("THU")
if weekday(matchdata,1)=6 then response.Write("FRI")
if weekday(matchdata,1)=7 then response.Write("SAT")
%> </td>
                <td align="center" bgcolor="#f5f5f5"><%=rs("hconfirm")%></td>
                <td height="35" align="center" bgcolor="#f5f5f5"><%=rs("htname")%> <br></td>
                <td bgcolor=<%=rs("hcolor")%>>&nbsp;</td>
                <td align="center" bgcolor="#FFFFFF">-</td>
                <td bgcolor=<%=rs("acolor")%>>&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5"><%=rs("atname")%></td>
                <td align="center" bgcolor="#f5f5f5"><%=rs("aconfirm")%></td>
                <td align="center" bgcolor="#f5f5f5"><%=rs("fname")%></td>
                <td align="center" bgcolor="#f5f5f5"> <a href="match_r_c.asp?mid=<%=rs("mid")%>"><%="ADD RESULT"%></a> </td>
              </tr>
              <%
end if
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
                <td height="35" colspan="14" align="center" bgcolor="#f5f5f5"><strong>ALL 
                  REPORT IS REFER</strong></td>
              </tr>
<%
end if
%>
              <tr bgcolor="#FFFFFF">
                <td height="50" colspan="14" align="center">&nbsp;</td>
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
