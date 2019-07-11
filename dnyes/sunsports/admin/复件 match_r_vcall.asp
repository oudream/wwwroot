<SCRIPT LANGUAGE="JavaScript">
<!--

function gettarg(targ,selObj,restore){ //v3.0
  eval(targ+".location='match_r_vcall.asp?cid="+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}

function getcidweek(targ,cid,selObj,restore){ //v3.0
  eval(targ+".location='match_r_vcall.asp?cid="+cid+"&mweek="+selObj.options[selObj.selectedIndex].value+"'");
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
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><table width="100%" cellpadding=0 cellspacing=0>
        <tbody>
          <tr> 
            <td width="50%" align="right"><b><font color=#000066 
                        face="Verdana, Arial, Helvetica, sans-serif">Select the 
              league the match in:</font></b></td>
            <td width=50% align="left"> <select name="cid" class="display_dropdown" id="cid" onchange="gettarg('self',this,0)">
                <option selected value="">SELECT ... </option>
                <%
sql="select * from tournament" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if rs.bof or rs.eof then response.Redirect("error.asp?err=1")
do while not rs.eof
%>
                <option value=<%=rs("cid")%>><%=rs("name")%> </option>
                <%
rs.movenext
loop
rs.close
set rs=nothing
%>
              </select></td>
          </tr>
        </tbody>
      </table></td>
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
            <table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
              <tr bgcolor="#FFFFFF"> 
                <td height="25" colspan="13"><b> <%=cname%> -- Manager have not refered report</b></td>
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
              <tr bgcolor="#FFFFFF"> 
                <td colspan="13"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="34%"><b>Week : <%=mweek%></b></td>
                      <td width="66%" align="right"><table cellpadding=0 cellspacing=0 width=350>
                          <tbody>
                            <tr> 
                              <td width=234 align="right"><b><font color=#000066 
                        face="Verdana, Arial, Helvetica, sans-serif">Display Results 
                                For Week:</font></b></td>
                              <td width=114 align="right"> <select name="mweek" id="mweek" class="display_dropdown" onchange="getcidweek('self',<%=cid%>,this,0)">
                                  <option selected>==Select ...==</option>
<%
dim amweek()
mwsql="select distinct mweek from match where cid="&cid&" order by mweek desc" 
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
                <td width="49" align="center" bgcolor="#FFFFFF"><b>Date</b></td>
                <td width="49" align="center" bgcolor="#FFFFFF"><b>Time</b></td>
                <td width="34" align="center" bgcolor="#FFFFFF"><b>Day</b></td>
                <td width="63" align="center" bgcolor="#FFFFFF"><b>Home Confirm</b></td>
                <td width="168" align="center" bgcolor="#FFFFFF"><b>Home Team</b></td>
                <td width="19" align="center" bgcolor="#FFFFFF"><b>CL</b></td>
                <td width="18" align="center" bgcolor="#FFFFFF"><b>VS</b></td>
                <td width="18" bgcolor="#FFFFFF"><b>CL</b></td>
                <td width="172" align="center" bgcolor="#FFFFFF"><b>Away Team</b></td>
                <td width="63" align="center" bgcolor="#FFFFFF"><b>Away Confirm</b></td>
                <td width="120" align="center" bgcolor="#FFFFFF"><b>Location</b></td>
                <td width="177" align="center" bgcolor="#FFFFFF"><strong>operation</strong></td>
              </tr>
<%
sql="select * from match where mweek="&mweek&" and cid="&cid&" order by myear,mmonth,mday,mhour,mminute desc" 
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.Open sql,conn,1,1
%>
<%
do while not rs.eof
if isnull(rs("matched")) then
%>
<%
matchdata=cstr(rs("mmonth"))&"-"&cstr(rs("mday"))&"-"&cstr(rs("myear"))
matchtime=cstr(rs("mhour"))&":"&cstr(rs("mminute"))
%>
<%
if DateDiff("d", date, cdate(matchdata))<0 then
%>
<%
ssql="select * from scorer where mid="&rs("mid")
Set srs=Server.CreateObject("ADODB.RecordSet") 
srs.Open ssql,conn,1,1
if not ( srs.eof or srs.bof ) then
do while not srs.eof
if srs("tid")=rs("htid") then hadded="yes"
if srs("tid")=rs("atid") then aadded="yes"
srs.movenext
loop
end if
srs.close
set srs=nothing
if not ( hadded="yes" and aadded="yes" ) then
%>
              <tr> 
                <td height="35" align="center" bgcolor="#f5f5f5"><%=matchdata%> </td>
                <td align="center" bgcolor="#f5f5f5"><%=matchtime%></td>
                <td align="center" bgcolor="#f5f5f5"><%=weekdayname(weekday(matchdata))%></td>
                <td align="center" bgcolor="#f5f5f5"><%=rs("hconfirm")%></td>
                <td height="35" align="center" bgcolor="#f5f5f5">
				<%=rs("htname")%><br><font color="#FF0000">
<%
if hadded="yes" then
response.Write("have reported")
else
response.Write("not reported")
end if
%>
				</font></td>
                <td bgcolor=<%=rs("hcolor")%>>&nbsp;</td>
                <td align="center" bgcolor="#FFFFFF">-</td>
                <td bgcolor=<%=rs("acolor")%>>&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5">
				<%=rs("atname")%><br><font color="#FF0000">
<%
if aadded="yes" then
response.Write("have reported")
else
response.Write("not reported")
end if
%>
				</font></td>
                <td align="center" bgcolor="#f5f5f5"><%=rs("aconfirm")%></td>
                <td align="center" bgcolor="#f5f5f5"><%=rs("fname")%></td>
                <td align="center" bgcolor="#f5f5f5">
				<a href="match_r_c1.asp?mid=<%=rs("mid")%>&htid=<%=rs("htid")%>&atid=<%=rs("atid")%>">Also to ADD Result</a>
				</td>
              </tr>
<%
end if
end if
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
