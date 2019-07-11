<!--#include file="adminconn.asp" -->


<%

if request("option")="update" then

hsnum=request("hsnum")
hcnum=request("hcnum")
hmmid=request("hmid")
htid=request("htid")
htname=request("htname")

asnum=request("asnum")
acnum=request("acnum")
mmid=request("mid")
tid=request("tid")
tname=request("tname")

conn.execute "delete * from scorer where mid="&hmmid&" and tid="&htid
conn.execute "delete * from card where mid="&hmmid&" and tid="&htid

if hsnum>0 then
dim asi()
dim aspid()
redim asi(hsnum-1)
redim aspid(hsnum-1)
for i=0 to ubound(asi)

aspid(i)=request("hspid"&cstr(i))
asi(i)=request("hs"&cstr(i))

psql="select * from player where pid=" & aspid(i)
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1 
pname=prs("name")
jersey=prs("jersey")
prs.close
set prs=nothing
conn.Execute("insert into scorer (mid,tid,tname,pid,pname,jersey,score) values ("&hmmid&","&htid&",'"&htname&"',"&aspid(i)&",'"&pname&"',"&jersey&","&asi(i)&")")

next
end if




if hcnum>0 then
dim adi()
dim adpid()
dim aci()
redim adi(hcnum-1)
redim adpid(hcnum-1)
redim aci(hcnum-1)
for i=0 to ubound(adpid)
adpid(i)=request("hdpid"&cstr(i))
aci(i)=request("hc"&cstr(i))
adi(i)=request("hd"&cstr(i))

psql="select * from player where pid=" & adpid(i)
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1 
pname=prs("name")
jersey=prs("jersey")
prs.close
set prs=nothing
if aci(i)="yellow" then conn.Execute("insert into card (mid,tid,tname,pid,pname,jersey,yellow,red) values ("&hmmid&","&htid&",'"&htname&"',"&adpid(i)&",'"&pname&"',"&jersey&","&adi(i)&",0)")
if aci(i)="red" then conn.Execute("insert into card (mid,tid,tname,pid,pname,jersey,red,yellow) values ("&hmmid&","&htid&",'"&htname&"',"&adpid(i)&",'"&pname&"',"&jersey&","&adi(i)&",0)")
next
end if




conn.execute "delete * from scorer where mid="&mmid&" and tid="&tid
conn.execute "delete * from card where mid="&mmid&" and tid="&tid

if asnum>0 then
redim asi(asnum-1)
redim aspid(asnum-1)
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
conn.Execute("insert into rresult (mid,tid,hscore,ascore,refer) values ("&mmid&","&tid&","&hhacountscore&","&aacountscore&",10)")
else
conn.Execute( "update rresult set hscore="&hhacountscore&", ascore="&aacountscore&" where mid="&mmid&" and tid="&tid )
end if
trs.close
set trs=nothing

tsql="select * from rresult where mid="&mmid&" and tid="&htid
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
if trs.eof or trs.bof then
conn.Execute("insert into rresult (mid,tid,hscore,ascore,refer) values ("&mmid&","&htid&","&hhacountscore&","&aacountscore&",10)")
else
conn.Execute( "update rresult set hscore="&hhacountscore&", ascore="&aacountscore&" where mid="&mmid&" and tid="&htid )
end if
trs.close
set trs=nothing

response.Redirect("match_r_e.asp?option=editresultissucc&mid="&mmid)
end if
%>
