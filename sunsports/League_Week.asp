<!--#include file="conn/conn.asp" -->
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
				  