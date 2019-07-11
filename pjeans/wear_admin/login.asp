<!--#include file="wefafs63f6asdfsdf6sdf4a.asp" -->
<%
if request("option")="logout" then
            session("user_uid")=""
            session("user_pwd")=""
			session("user_name")=""
            session("user_adminlevel")=0
            session("user_email")=""
            session("user_tel")=""
            session("user_zip")=""
			session("user_pid")=""
			session("manager_pid")=""
			session("manager_tid")=""
			session("manager_tname")=""
response.redirect("index.asp")
response.end
end if
%>
<%
uid=trim(request("uid"))
pwd=trim(request("pwd"))

sql="select * from admin_users where uid='"&uid&"'"
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
if rs.eof or rs.bof then
%>
<SCRIPT LANGUAGE=JavaScript>
alert(' Please enter a valid Userid . ');
window.history.go( -1 );
</SCRIPT>
<%
response.end
end if

if pwd<>rs("pwd") then
%>
<SCRIPT LANGUAGE=JavaScript>
alert(' Please enter a valid Userid and Password . ');
window.history.go( -1 );
</SCRIPT>
<%
response.end
end if

session("user_pid")=trim(rs("pid"))
session("user_uid")=trim(rs("uid"))
session("user_pwd")=trim(rs("pwd"))
session("user_name")=trim(rs("name"))
session("user_adminlevel")=rs("adminlevel")
session("user_email")=trim(rs("email"))
session("user_tel")=trim(rs("tel"))
session("user_zip")=trim(rs("zip"))
rs.close
set rs=nothing

%>
<%
response.redirect("adminlogin.asp")
response.end
%>
