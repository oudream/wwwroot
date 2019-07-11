<!--#include file="conn/conn.asp" -->
<%
rrstr="no"
dim rrhscore(1)
dim rrascore(1)

rsql="select * from rresult where mid=18"
Set rrs=Server.CreateObject("ADODB.RecordSet")
rrs.Open rsql,conn,1,1
zi=0
do while not rrs.eof
rrhscore(zi)=rrs("hscore")
rrascore(zi)=rrs("ascore")
zi=zi+1
rrs.movenext
loop
rrs.close
set rrs=nothing
if isempty(rrhscore(0)) or isempty(rrhscore(1)) or isempty(rrascore(0)) or isempty(rrascore(1)) then
rrstr="123"
else
if rrhscore(0)=rrascore(0) and rrhscore(1)=rrascore(1) then rrstr="yes"
end if
%>

<%=rrstr%>

<br>
<br>
<br>
<br>
<%=rrhscore(0)%>
<br>
<br>
<br>
<%=rrhscore(1)%>
<br>
<br>
<br>
<br>
<%=rrascore(0)%>
<br>
<br>
<br>
<%=rrascore(1)%>
<br>
<br>
<br>
<br>
fsdaf