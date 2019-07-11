<!--#include file="adminconn.asp" -->


<%
mmid=request("mid")
htid=request("htid")
tid=request("tid")

hhacountscore=0
sql="select * from scorer where mid="&mmid&" and tid="&htid
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1
do while not rs.eof 
hhacountscore=hhacountscore + rs("score")
rs.movenext
loop
rs.close

aacountscore=0
sql="select * from scorer where mid="&mmid&" and tid="&tid
rs.Open sql,conn,1,1
do while not rs.eof 
aacountscore=aacountscore + rs("score")
rs.movenext
loop
rs.close
set rs=nothing

tsql="select * from rresult where mid="&mmid&" and tid="&tid
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
if trs.eof or trs.bof then
conn.Execute("insert into rresult (mid,tid,hscore,ascore,refer) values ("&mmid&","&tid&",310,10,10)")
else
conn.Execute( "update rresult set hscore="&hhacountscore&", ascore=20 where mid="&mmid&" and tid="&tid )
end if
trs.close
set trs=nothing



response.Write(aacountscore&"<br><br>")
response.Write(hhacountscore&"<br><br>")
response.Write(mmid&"<br><br>")
response.Write(tid&"<br><br>")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Create match is complete. ');</SCRIPT>")

%>
