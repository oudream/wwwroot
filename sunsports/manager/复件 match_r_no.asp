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
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td align="center" class="header">ALL MATCH</td>
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
msql="select mweek from match where htid="&session("manager_tid")&" or atid="&session("manager_tid")&" order by mweek desc" 
Set mrs=Server.CreateObject("ADODB.RecordSet") 
mrs.Open msql,conn,1,1 
%> 
            <table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
              <tr bgcolor="#FFFFFF"> 
                <td height="25" colspan="13"><b> Your team -- Fixture and Result 
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
                <td colspan="13"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="34%"><b>Week : <%=mweek%></b></td>
                      <td width="66%" align="right"><table cellpadding=0 cellspacing=0 width=350>
                          <tbody>
                            <tr> 
                              <td width=234 align="right"><b><font color=#000066 
                        face="Verdana, Arial, Helvetica, sans-serif">Display Results For Week:</font></b></td>
                              <td width=114 align="right"> <select name="mweek" id="mweek" class="display_dropdown" onchange="getweek('self',this,0)">
                                  <option selected value=<%=amweek%>>SELECT ..</option>
                                  <%
dim amweek()
mwsql="select distinct mweek from match where htid="&session("manager_tid")&" or atid="&session("manager_tid")&" order by mweek desc" 
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
                                  <option value=<%=amweek(i)%>>==<%=amweek(i)%>==<%=minday%>/<%=minmonth%>/<%=minyear%>--<%=maxday%>/<%=maxmonth%>/<%=maxyear%></option>
                                  <%
next
end if
erase amweek
%>
                                </select> </td>
                            </tr>
                          </tbody>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td width="180" align="center" bgcolor="#FFFFFF"><strong>Tournament</strong></td>
                <td width="157" align="center" bgcolor="#FFFFFF"><b>Date</b></td>
                <td width="34" align="center" bgcolor="#FFFFFF"><b>Time</b></td>
                <td width="56" align="center" bgcolor="#FFFFFF"><b>Day</b></td>
                <td width="60" align="center" bgcolor="#FFFFFF"><b>Home Confirm</b></td>
                <td width="66" align="center" bgcolor="#FFFFFF"><b>Home Team</b></td>
                <td width="25" align="center" bgcolor="#FFFFFF"><b>CL</b></td>
                <td width="38" align="center" bgcolor="#FFFFFF"><b>VS</b></td>
                <td width="25" bgcolor="#FFFFFF"><b>CL</b></td>
                <td width="80" align="center" bgcolor="#FFFFFF"><b>Away Team</b></td>
                <td width="88" align="center" bgcolor="#FFFFFF"><b>Away Confirm</b></td>
                <td width="141" align="center" bgcolor="#FFFFFF"><b>Location</b></td>
                <td width="180" align="center" bgcolor="#FFFFFF"><strong>Operation</strong></td>
              </tr>
              <%
sql="select * from match where mweek="&mweek&" and ( htid="&session("manager_tid")&" or atid="&session("manager_tid")&") order by myear,mmonth,mday,mhour,mminute desc" 
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.Open sql,conn,1,1
%>
              <%
do while not rs.eof
if ( rs("htid")=session("manager_tid") and isnull(rs("matched")) ) or ( rs("atid")=session("manager_tid") and isnull(rs("matched")) ) then
if rs("htid")=session("manager_tid") then ha="h"
if rs("atid")=session("manager_tid") then ha="a"
matchdata=cstr(rs("mmonth"))&"-"&cstr(rs("mday"))&"-"&cstr(rs("myear"))
matchtime=cstr(rs("mhour"))&":"&cstr(rs("mminute"))
%>
              <tr> 
                <td align="center" bgcolor="#f5f5f5"><%=rs("cname")%>&nbsp;</td>
                <td height="35" align="center" bgcolor="#f5f5f5"><%=matchdata%> 
                </td>
                <td align="center" bgcolor="#f5f5f5"><%=matchtime%></td>
                <td align="center" bgcolor="#f5f5f5"><%=weekdayname(weekday(matchdata))%></td>
                <td align="center" bgcolor="#f5f5f5"><%=rs("hconfirm")%></td>
                <td height="35" align="center" bgcolor="#f5f5f5"><%=rs("htname")%> 
                  <br></td>
                <td bgcolor=<%=rs("hcolor")%>>&nbsp;</td>
                <td align="center" bgcolor="#FFFFFF">- 
                  <!-- --------------------------------------------------------------------------------------------------
													assign to "hsumscorer or asumscorer"
---------------------------------------------------------------------------------------------------- -->
                </td>
                <td bgcolor=<%=rs("acolor")%>>&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5"><%=rs("atname")%></td>
                <td align="center" bgcolor="#f5f5f5"><%=rs("aconfirm")%></td>
                <td align="center" bgcolor="#f5f5f5"><%=rs("fname")%></td>
                <td align="center" bgcolor="#f5f5f5"> 
                  <%
ssql="select * from rscorer where mid="&rs("mid")&" and tid="&session("manager_tid")
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
%>
                  <%
if tfedit="yes" then 
%>
                  <a href="match_r_v.asp?mid=<%=rs("mid")%>"><%="VIEW&EDIT"%></a> 
                  <%
end if
%>
                  <%
if tfedit="no" then 
%>
                  <a href="match_r_c.asp?mid=<%=rs("mid")%>"><%="ADD RESULT"%></a> 
                  <%
end if
%>
                </td>
              </tr>
              <%
end if
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
      </table></td>
    <td>&nbsp;</td>
  </tr>
</table>
