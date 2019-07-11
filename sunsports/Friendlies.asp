<SCRIPT LANGUAGE="JavaScript">
<!--
function getcidweek(targ,selObj,restore){ //v3.0
  eval(targ+".location='Friendlies.asp?mweek="+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}

//-->

</SCRIPT>
<!--#include file="Top_sun.asp"-->
<table width="776" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
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
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="League.asp"><b><font color="#1F4A71">LEAGUE MATCHES</font></b></a><font color="#1F4A71">+ 
                  +</font></td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="Cup.asp"><b><font color="#1F4A71">CUP MATCHES </font></b></a><font color="#1F4A71">&nbsp;+ 
                  +</font></td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="Friendlies.asp"><b><font color="#1F4A71">FRIENDLY MATCHES 
                  </font></b></a></td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="TeamProfile.asp"><b><font color="#1F4A71">TEAMPORFILE</font></b></a> 
                  <font color="#1F4A71">&nbsp;+ +</font> </td>
              </tr>
            </table>
            <!-- ========================================================================================================		  

													cup table ============stop

 ======================================================================================================== -->
          </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table></td>
    <td width="579" valign="top"><br>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="20">&nbsp;</td>
          <td>
		  
		  
		  
<%
msql="select mweek from match where cid=0 order by mweek desc" 
Set mrs=Server.CreateObject("ADODB.RecordSet") 
mrs.Open msql,conn,1,1 
%> 
            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#6184A3">
              <tr> 
                <td height="25" colspan="12" bgcolor="#FFFFFF"><b>Recent Friendlies Matchs Fixture</b></td>
              </tr>
              <%
if mrs.eof or mrs.bof then
response.Write("<tr bgcolor=#FFFFFF><td height=25 colspan=11><a href=friendlies_c.asp> <b>ADD FRIENDLIES</b></a></td></tr></table>")
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
                <td colspan="12" bgcolor="#FFFFFF"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="34%"><b>Week : <%=mweek%></b></td>
                      <td width="66%" align="right"><table cellpadding=0 cellspacing=0 width=350>
                          <tbody>
                            <tr> 
                              <td width=234 align="right"><b><font color=#000066 
                        face="Verdana, Arial, Helvetica, sans-serif">Display Results 
                                For Week:</font></b></td>
                              <td width=114 align="right"> <select name="mweek" id="mweek" class="display_dropdown" onchange="getcidweek('self',this,0)">
                                  <option value="" selected>SELECT WEEK ...</option>
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
    <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                          onmouseover="this.style.backgroundColor='#BECFDF';"> 
                <td width="60" align="center" ><b>Date</b></td>
                <td width="35" align="center" ><b>Time</b></td>
                <td width="35" align="center" ><b>Day</b></td>
                <td width="45" align="center" ><b>Home Confirm</b></td>
                <td align="center" ><b>Home Team</b></td>
                <td width="13" align="center" ><b>CL</b></td>
                <td width="30" align="center" ><b>VS</b></td>
                <td width="13" ><b>CL</b></td>
                <td align="center" ><b>Away Team</b></td>
                <td width="45" align="center" ><b>Away Confirm</b></td>
                <td width="91" align="center" ><b>Location</b></td>
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
    <tr bgcolor="#FFFFFF" onmouseout="this.style.backgroundColor='#FFFFFF';" 
                          onmouseover="this.style.backgroundColor='#BECFDF';"> 
                <td height="28" align="center"><%=matchdata%>&nbsp;</td>
                <td align="center"><%=matchtime%>&nbsp;</td>
                <td align="center"> <%
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
                <td align="center"><%=rs("htname")%>&nbsp;</td>
                <td bgcolor=<%=rs("hcolor")%>>&nbsp;</td>
                <td align="center">-</td>
                <td bgcolor=<%=rs("acolor")%>>&nbsp;</td>
                <td align="center"><%=rs("atname")%>&nbsp;</td>
                <td align="center"><%=rs("aconfirm")%>&nbsp;</td>
                <td align="center"><%=rs("fname")%>&nbsp;</td>
              </tr>
              <%
rs.movenext
loop
rs.close
set rs=nothing
%>
            </table>
		  
		  
		  
		  </td>
        </tr>
      </table> 
      
    </td>
        </tr>
      </table> 
<!--#include file="Copyright.asp"-->