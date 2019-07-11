<!--#include file="adminconn.asp" -->
<%
tid=session("manager_tid")
mmid=request("mid")
if tid="" and mmid="" then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Please select the team you want to add scorer or discipline at first ! ');</SCRIPT>")
response.End()
end if
%>
<%
snum=request("snum")
dnum=request("dnum")
hscore=request("hscore")
ascore=request("ascore")
refer=request("refer")
%>
<%
if request("option")="update" then

conn.execute "delete * from scorer where mid="&mmid&" and tid="&tid

conn.execute "delete * from card where mid="&mmid&" and tid="&tid

if snum>0 then
dim asi()
dim aspid()
redim asi(snum-1)
redim aspid(snum-1)
for i=0 to ubound(asi)
aspid(i)=request("spid"&cstr(i))
asi(i)=request("s"&cstr(i))
ssql="select * from scorer where mid="&mmid&" and pid="&aspid(i)
Set srs=Server.CreateObject("ADODB.RecordSet") 
srs.Open ssql,conn,1,1 

if not ( srs.eof or srs.bof ) then
tfstr="yes"
else
tfstr="no"
end if

srs.close
set srs=nothing

if tfstr="yes" then 
conn.Execute("update scorer set score="&asi(i)&" where mid="&mmid&" and pid="&aspid(i))
end if

if tfstr="no" then
psql="select * from player where pid=" & aspid(i)
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1 
pname=prs("name")
jersey=prs("jersey")
prs.close
set prs=nothing
conn.Execute("insert into scorer (mid,tid,tname,pid,pname,jersey,score) values ("&mmid&","&session("manager_tid")&",'"&session("manager_tname")&"',"&aspid(i)&",'"&pname&"',"&jersey&","&asi(i)&")")
end if
next
erase aspid
erase asi
end if




if dnum>0 then
dim adi()
dim adpid()
dim aci()
redim adi(dnum-1)
redim adpid(dnum-1)
redim aci(dnum-1)
for i=0 to ubound(adpid)
adpid(i)=request("dpid"&cstr(i))
aci(i)=request("c"&cstr(i))
adi(i)=request("d"&cstr(i))
ssql="select * from card where mid="&mmid&" and pid="&adpid(i)
Set srs=Server.CreateObject("ADODB.RecordSet") 
srs.Open ssql,conn,1,1 

if not ( srs.eof or srs.bof ) then
tfstr="yes"
else
tfstr="no"
end if

srs.close
set srs=nothing

if tfstr="yes" then
if aci(i)="yellow" then conn.Execute("update card set yellow="&adi(i)&" , red=0 where mid="&mmid&" and pid="&adpid(i))
if aci(i)="red" then conn.Execute("update card set red="&adi(i)&" , yellow=0 where mid="&mmid&" and pid="&adpid(i))
end if

if tfstr="no" then
psql="select * from player where pid=" & adpid(i)
Set prs=Server.CreateObject("ADODB.RecordSet") 
prs.Open psql,conn,1,1 
pname=prs("name")
jersey=prs("jersey")
prs.close
set prs=nothing
if aci(i)="yellow" then conn.Execute("insert into card (mid,tid,tname,pid,pname,jersey,yellow,red) values ("&mmid&","&session("manager_tid")&",'"&session("manager_tname")&"',"&adpid(i)&",'"&pname&"',"&jersey&","&adi(i)&",0)")
if aci(i)="red" then conn.Execute("insert into card (mid,tid,tname,pid,pname,jersey,red,yellow) values ("&mmid&","&session("manager_tid")&",'"&session("manager_tname")&"',"&adpid(i)&",'"&pname&"',"&jersey&","&adi(i)&",0)")
end if
next
erase adpid
erase aci
erase adi
end if

acountscore=0
sql="select * from scorer where mid="&mmid&" and tid="&tid
rs.Open sql,conn,1,1
do while not rs.eof 
acountscore=acountscore + rs("score")
rs.movenext
loop
rs.close
set rs=nothing

if request("homeaway")="h" then hscore=acountscore
if request("homeaway")="a" then ascore=acountscore

tsql="select * from rresult where mid="&mmid&" and tid="&tid
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
if trs.eof or trs.bof then
conn.Execute("insert into rresult (mid,tid,hscore,ascore,refer) values ("&mmid&","&tid&","&hscore&","&ascore&","&refer&")")
else
conn.Execute( "update rresult set hscore="&hscore&", ascore="&ascore&" where mid="&mmid&" and tid="&tid )
end if
trs.close
set trs=nothing

response.Redirect("match_r_e.asp?option=editresultissucc&mid="&mmid)
end if
%>
