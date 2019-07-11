<SCRIPT LANGUAGE="JavaScript">
<!--

function gettarg(targ,selObj,restore){ //v3.0
  eval(targ+".location='match_r_vma.asp?cid="+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}

function getcidweek(targ,cid,selObj,restore){ //v3.0
  eval(targ+".location='match_r_vma.asp?cid="+cid+"&mweek="+selObj.options[selObj.selectedIndex].value+"'");
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
                <td height="25" colspan="13"><b> <font color="#FF0000"><%=cname%></font> -- Manager have not refered report</b></td>
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
                <td colspan="13">&nbsp; </td>
              </tr>
              <tr> 
                <td width="49" align="center" bgcolor="#FFFFFF"><b>Week</b></td>
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
sql="select * from match where cid="&cid&" order by myear,mmonth,mday,mhour,mminute desc" 
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.Open sql,conn,1,1
%>
<%
do while not rs.eof
if rs("matched")="no" then
%>
<%
matchdata=cstr(rs("mmonth"))&"/"&cstr(rs("mday"))&"/"&cstr(rs("myear"))
matchtime=cstr(rs("mhour"))&":"&cstr(rs("mminute"))
%>
<%
if DateDiff("d", date, cdate(matchdata))<=0 then
%>
<%
hadded="no"
aadded="no"
ssql="select * from rresult where mid="&rs("mid")
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
if hadded="yes" and aadded="yes" then
%>
              <tr> 
                <td height="35" align="center" bgcolor="#f5f5f5"><%=rs("mweek")%>&nbsp;</td>
                <td height="35" align="center" bgcolor="#f5f5f5"><%=matchdata%>&nbsp;</td>
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
				&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5"><%=rs("hconfirm")%>&nbsp;</td>
                <td height="35" align="center" bgcolor="#f5f5f5">
				<%=rs("htname")%><br><font color="#FF0000">
<%
if hadded="yes" then
response.Write("have reported")
else
response.Write("not reported")
end if
%>
				</font>&nbsp;</td>
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
				</font>&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5"><%=rs("aconfirm")%>&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5"><%=rs("fname")%>&nbsp;</td>
                <td align="center" bgcolor="#f5f5f5">
				<a href="match_r_vh.asp?mid=<%=rs("mid")%>&htid=<%=rs("htid")%>&atid=<%=rs("atid")%>">Also to ADD Result</a>
				&nbsp;</td>
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
