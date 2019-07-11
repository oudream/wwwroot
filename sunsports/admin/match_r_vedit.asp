<SCRIPT LANGUAGE="JavaScript">
<!--

function gettarg(targ,selObj,restore){ //v3.0
  eval(targ+".location='match_r_vedit.asp?cid="+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}

function getcidweek(targ,cid,selObj,restore){ //v3.0
  eval(targ+".location='match_r_vedit.asp?cid="+cid+"&mweek="+selObj.options[selObj.selectedIndex].value+"'");
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
                <option selected value="">SELECT LEAGUE ... </option>
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
                <td height="25" colspan="13"><b>ALL SUMBITED REPORTS OF -- <font color="#FF0000"><%=cname%></font></b></td>
              </tr>
              <%
if mrs.eof or mrs.bof then
response.Write("<tr bgcolor=#FFFFFF><td height=25 colspan=11><strong>NO MATCH , YOU CAN : <a href=match_c.asp?cid="&cid&"> ADD MATCH</a></strong></td></tr></table>")
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
                                  <option selected>==SELECT WEEK ...==</option>
                                  <option>-----------------------------</option>
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
                <td width="39" align="center" bgcolor="#FFFFFF"><b>Date</b></td>
                <td width="39" align="center" bgcolor="#FFFFFF"><b>Time</b></td>
                <td width="29" align="center" bgcolor="#FFFFFF"><b>Day</b></td>
                <td width="30" align="center" bgcolor="#FFFFFF"><b>Home Confirm</b></td>
                <td width="115" align="center" bgcolor="#FFFFFF"><b>Home Team</b></td>
                <td width="15" align="center" bgcolor="#FFFFFF"><b>CL</b></td>
                <td width="30" align="center" bgcolor="#FFFFFF"><b>VS</b></td>
                <td width="15" bgcolor="#FFFFFF"><b>CL</b></td>
                <td width="118" align="center" bgcolor="#FFFFFF"><b>Away Team</b></td>
                <td width="30" align="center" bgcolor="#FFFFFF"><b>Away Confirm</b></td>
                <td width="127" align="center" bgcolor="#FFFFFF"><b>Location</b></td>
                <td width="82" align="center" bgcolor="#FFFFFF"><strong>operation</strong></td>
              </tr>
<%
sql="select * from match where mweek="&mweek&" and cid="&cid&" order by mweek desc,mday desc" 
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.Open sql,conn,1,1
%>
<%
dim rrhscore(1)
dim rrascore(1)

do while not rs.eof

hadded="no"
aadded="no"
rrstr="no"
ouifdata="no"
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

if rrstr="yes" and hadded="yes" and aadded="yes" then 

ouifdata="yes"
%>
<%
matchdata=cstr(rs("mmonth"))&"/"&cstr(rs("mday"))&"/"&cstr(rs("myear"))
matchtime=cstr(rs("mhour"))&":"&cstr(rs("mminute"))
%>
              <tr> 
                <td height="35" align="center" bgcolor="#f5f5f5"><%=matchdata%></td>
                <td align="center" bgcolor="#f5f5f5"><%=matchtime%>&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5">
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
                <td align="center" bgcolor="#f5f5f5"><%=rs("hconfirm")%></td>
                <td height="35" align="center" bgcolor="#f5f5f5">
				<%=rs("htname")%>&nbsp;</td>
                <td bgcolor=<%=rs("hcolor")%>>&nbsp;</td>
                <td align="center" bgcolor="#FFFFFF"><%=rrhscore(0)&" : "&rrascore(0)%></td>
                <td bgcolor=<%=rs("acolor")%>>&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5">
				<%=rs("atname")%></td>
                <td align="center" bgcolor="#f5f5f5"><%=rs("aconfirm")%>&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5"><%=rs("fname")%>&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5"> <a href="match_r_e.asp?mid=<%=rs("mid")%>">EDIT</a></td>
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
                <td height="35" colspan="13" align="center" bgcolor="#f5f5f5"><strong>NOT Submitted Reports</strong></td>
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