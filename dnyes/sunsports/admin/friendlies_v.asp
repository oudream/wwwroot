<SCRIPT LANGUAGE="JavaScript">
<!--
function getcidweek(targ,selObj,restore){ //v3.0
  eval(targ+".location='friendlies_v.asp?mweek="+selObj.options[selObj.selectedIndex].value+"'");
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
        Friendlies Fixture Edit/Delete<br>
        <br>
        </FONT></H2></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
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
if trim(request("option"))="del" then
conn.execute "delete * from match where mid=" & request("mid")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Del Match is complete. ');</SCRIPT>")
response.Redirect("friendlies_v.asp")
response.End()
end if
%>
<%
msql="select mweek from match where cid=0 order by mweek desc" 
Set mrs=Server.CreateObject("ADODB.RecordSet") 
mrs.Open msql,conn,1,1 
%> 
            <table width="100%" border="1" align="center" cellpadding="2" cellspacing="0" bordercolor="#CCCCCC">
              <tr bgcolor="#FFFFFF"> 
                <td height="25" colspan="12"><b> </b></td>
              </tr>
              <%
if mrs.eof or mrs.bof then
response.Write("<tr bgcolor=#FFFFFF><td height=25 colspan=11><a href=friendlies_c.asp> <b>Add Friendlies Fixture</b></a></td></tr></table>")
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
                <td colspan="12"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="34%"><b>Week : <%=mweek%></b></td>
                      <td width="66%" align="right"><table cellpadding=0 cellspacing=0 width=350>
                          <tbody>
                            <tr> 
                              <td width=234 align="right"><b><font color=#000066 
                        face="Verdana, Arial, Helvetica, sans-serif">Display Results 
                                For Week:</font></b></td>
                              <td width=114 align="right"> <select name="mweek" id="mweek" class="display_dropdown" onchange="getcidweek('self',this,0)">
                                  <option value="" selected>==SELECT WEEK ...===</option>
                                  <option value="">---------------------------</option>
                                  <%
dim amweek()
mwsql="select distinct mweek from match where cid=0 order by mweek desc" 
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
sql="select * from match where mweek="&mweek&" and cid=0 order by myear,mmonth,mday,mhour,mminute desc" 
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.Open sql,conn,1,1
%>
              <%
do while not rs.eof
matchdata=cstr(rs("mmonth"))&"/"&cstr(rs("mday"))&"/"&cstr(rs("myear"))
matchtime=cstr(rs("mhour"))&":"&cstr(rs("mminute"))
%>
              <tr> 
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
                <td align="center" bgcolor="#f5f5f5">
				<a href="friendlies_e.asp?mid=<%=rs("mid")%>"> Edit </a> &nbsp;&nbsp;
				<a href="friendlies_v.asp?mweek=<%=mweek%>&option=del&mid=<%=rs("mid")%>" onClick="return confirm(' Are you sure to del the friendlies ?')"> Del </a></td>
              </tr>
              <%
rs.movenext
loop
rs.close
set rs=nothing
%>
              <tr> 
                <td height="35" colspan="12" align="center" bgcolor="#f5f5f5"><a href=friendlies_c.asp> 
                  <b>Add Friendlies Fixture</b></a></td>
              </tr>
              <tr>
                <td height="50" colspan="12" align="center">&nbsp;</td>
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
