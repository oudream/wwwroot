<SCRIPT LANGUAGE="JavaScript">
<!--

function gettarg(targ,selObj,restore){ //v3.0
  eval(targ+".location='match_v.asp?cid="+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}

function getcidweek(targ,cid,selObj,restore){ //v3.0
  eval(targ+".location='match_v.asp?cid="+cid+"&mweek="+selObj.options[selObj.selectedIndex].value+"'");
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
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr>
    <td>&nbsp;</td>
    <td align="center"><H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066> <br>
        Fixture Edit/Delete<br>
        <br>
        </FONT></H2></td>
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
msql="select mweek from match" 
Set mrs=Server.CreateObject("ADODB.RecordSet") 
mrs.Open msql,conn,1,1 
%> 
            <table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
              <tr> 
                <td height="25" colspan="14" bgcolor="#FFFFFF"><b> <font color="#FF0000">All </font>-- 
                  Fixture and Result </b></td>
              </tr>
              <%
if mrs.eof or mrs.bof then
response.Write("<tr bgcolor=#FFFFFF><td height=25 colspan=11><a href=match_c.asp?cid="&cid&"> <b>ADD MATCH</b></a></td></tr></table>")
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
                <td colspan="14" bgcolor="#FFFFFF">&nbsp; </td>
              </tr>
              <tr> 
                <td width="60" align="center" bgcolor="#FFFFFF"><strong>Week</strong></td>
                <td width="60" align="center" bgcolor="#FFFFFF"><strong>Tournament</strong></td>
                <td width="80" align="center" bgcolor="#FFFFFF"><b>Date</b></td>
                <td width="30" align="center" bgcolor="#FFFFFF"><b>Time</b></td>
                <td width="30" align="center" bgcolor="#FFFFFF"><b>Day</b></td>
                <td width="30" align="center" bgcolor="#FFFFFF"><b>Home Confirm</b></td>
                <td width="100" align="center" bgcolor="#FFFFFF"><b>Home Team</b></td>
                <td width="15" align="center" bgcolor="#FFFFFF"><b>CL</b></td>
                <td width="30" align="center" bgcolor="#FFFFFF"><b>VS</b></td>
                <td width="15" bgcolor="#FFFFFF"><b>CL</b></td>
                <td width="100" align="center" bgcolor="#FFFFFF"><b>Away Team</b></td>
                <td width="30" align="center" bgcolor="#FFFFFF"><b>Away Confirm</b></td>
                <td width="100" align="center" bgcolor="#FFFFFF"><b>Location</b></td>
                <td width="60" align="center" bgcolor="#FFFFFF"><strong>Operation</strong></td>
              </tr>
              <%
dim rrhscore(1)
dim rrascore(1)
ouiffixture="no"
sql="select * from match order by mweek desc,mday desc" 
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.Open sql,conn,1,1
%>
              <%
do while not rs.eof
matchdata=cstr(rs("mmonth"))&"/"&cstr(rs("mday"))&"/"&cstr(rs("myear"))
matchtime=cstr(rs("mhour"))&":"&cstr(rs("mminute"))
if DateDiff("d", date, cdate(matchdata))>=0 then
ouiffixture="yes"
%>
              <tr> 
                <td align="center" bgcolor="#f5f5f5"><%=rs("mweek")%></td>
                <td align="center" bgcolor="#f5f5f5"><%=rs("cname")%></td>
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
%> &nbsp;</td>
                <td align="center" bgcolor="#f5f5f5"><%=rs("hconfirm")%>&nbsp;</td>
                <td height="35" align="center" bgcolor="#f5f5f5"><%=rs("htname")%>&nbsp;</td>
                <td bgcolor=<%=rs("hcolor")%>>&nbsp;</td>
                <td align="center" bgcolor="#FFFFFF">-</td>
                <td bgcolor=<%=rs("acolor")%>>&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5"><%=rs("atname")%>&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5"><%=rs("aconfirm")%>&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5"><%=rs("fname")%>&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5"><a href="match_e.asp?mid=<%=rs("mid")%>">Edit/Del</a></td>
              </tr>
              <%
end if
rs.movenext
loop
rs.close
set rs=nothing
%>
              <%
if ouiffixture="no" then
%>
              <tr> 
                <td height="35" colspan="14" align="center" bgcolor="#f5f5f5"><table width="100%" border="0" cellspacing="2" cellpadding="2">
                    <tr>
                      <td width="150"><strong>NOT FIXTURE , YOU CAN:</strong></td>
                      <td width="76%"><strong><a href="match_c.asp">ADD FIXTURE</a></strong></td>
                    </tr>
                  </table>
                  </td>
              </tr>
<%end if%>
              <tr>
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
