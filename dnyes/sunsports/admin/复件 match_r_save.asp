<!--#include file="adminconn.asp" -->
<%
if request("option")="add" then

hsnum=request("hsnum")
acnum=request("acnum")
mmid=request("mid")
tid=request("tid")
tname=request("tname")

conn.execute "delete * from scorer where mid="&mmid&" and tid="&tid
conn.execute "delete * from card where mid="&mmid&" and tid="&tid

if hsnum>0 then
dim asi()
dim aspid()
redim asi(hsnum-1)
redim aspid(hsnum-1)
for i=0 to ubound(asi)

aspid(i)=request("spid"&cstr(i))
asi(i)=request("s"&cstr(i))

psql="select * from player where pid=" & aspid(i)
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1 
pname=prs("name")
jersey=prs("jersey")
prs.close
set prs=nothing
conn.Execute("insert into scorer (mid,tid,tname,pid,pname,jersey,score) values ("&mmid&","&tid&",'"&tname&"',"&aspid(i)&",'"&pname&"',"&jersey&","&asi(i)&")")

next
erase aspid
erase asi
end if




if acnum>0 then
dim adi()
dim adpid()
dim aci()
redim adi(acnum-1)
redim adpid(acnum-1)
redim aci(acnum-1)
for i=0 to ubound(adpid)
adpid(i)=request("dpid"&cstr(i))
aci(i)=request("c"&cstr(i))
adi(i)=request("d"&cstr(i))

psql="select * from player where pid=" & adpid(i)
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1 
pname=prs("name")
jersey=prs("jersey")
prs.close
set prs=nothing
if aci(i)="yellow" then conn.Execute("insert into card (mid,tid,tname,pid,pname,jersey,yellow,red) values ("&mmid&","&tid&",'"&tname&"',"&adpid(i)&",'"&pname&"',"&jersey&","&adi(i)&",0)")
if aci(i)="red" then conn.Execute("insert into card (mid,tid,tname,pid,pname,jersey,red,yellow) values ("&mmid&","&tid&",'"&tname&"',"&adpid(i)&",'"&pname&"',"&jersey&","&adi(i)&",0)")
next
erase adpid
erase aci
erase adi
end if

tsql="select * from rresult where mid="&mmid&" and tid="&tid
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
if trs.eof or trs.bof then
conn.Execute("insert into rresult (mid,tid,hscore,ascore,refer) values ("&mmid&","&tid&","&hscore&","&ascore&","&refer&")")
else
conn.Execute( "update rresult set hscore="&request("hscore")&", ascore="&request("ascore")&" where mid="&mmid&" and tid="&tid )
end if
trs.close
set trs=nothing

acountscore=0
sql="select * from scorer where mid="&mmid&" and tid="&tid
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1
do while not rs.eof 
acountscore=acountscore + rs("score")
rs.movenext
loop
rs.close
set rs=nothing

if request("tstr")="h" then
if clng(acountscore)<>clng(request("hscore")) then response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' h ');</SCRIPT>")
end if
if request("tstr")="a" then
if clng(acountscore)<>clng(request("ascore")) then response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' a ');</SCRIPT>")
end if
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Create match is complete. ');</SCRIPT>")
end if
%>
