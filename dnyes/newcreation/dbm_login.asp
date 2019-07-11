<!--#include file="dbm_wellcome_conn.asp" -->
<%
            session("user_id")=0
			session("user_uid")=""
			session("user_allname")=""
            session("user_adminlevel")=0
			session("guest_allname")=""
if request("option")="logout" then
            session("user_id")=0
			session("user_uid")=""
			session("user_allname")=""
            session("user_adminlevel")=0
			session("guest_allname")=""
response.write "<script language='javascript'>window.close();</script>"
response.end
end if
%>
<%
uid=trim(request("uid"))
pwd=trim(request("pwd"))

if request("user_adminlevel")=9 then
sql="select * from dbm_guest where gid='"&uid&"'"
else
sql="select * from dbm_user where uid='"&uid&"'"
end if
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
if rs.eof or rs.bof then
response.Redirect("dbm_index.asp?err=uid_err")
end if

if pwd<>rs("pwd") then
response.Redirect("dbm_index.asp?err=pwd_err")
end if


if request("user_adminlevel")=9 then
session("user_id")=rs("guest_id")
session("user_uid")=rs("gid")
session("user_allname")=rs("allname")
session("user_adminlevel")=9
else
session("user_id")=rs("user_id")
session("user_uid")=rs("uid")
session("user_allname")=rs("allname")
session("user_adminlevel")=rs("adminlevel")
end if
rs.close
set rs=nothing

%>
<%

if session("user_id")<>0 then
tsql="select allname from dbm_guest where guest_id in ( select guest_id from dbm_ug where user_id="&session("user_id")&" )"
Set trs=Server.CreateObject("ADODB.RecordSet") 
trs.Open tsql,conn,1,1 
do while not trs.eof
session("guest_allname") = session("guest_allname") & "," & trs("allname")
trs.movenext
loop
trs.close
set trs=nothing
else session("guest_allname")=""
end if

response.redirect("dbm_index.asp")
response.end
%>
