<!--#include file="wefafs63f6asdfsdf6sdf4a.asp" -->
<%
if request("option")="logout" then
            session("user_id")=0
			session("user_name")=""
            session("user_adminlevel")=0
response.redirect("index.asp")
response.end
end if
%>
<%
uid=trim(request("uid"))
pwd=trim(request("pwd"))

sql="select * from user where uid='"&uid&"'"
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

session("user_id")=rs("id")
session("user_name")=rs("uname")
session("user_adminlevel")=rs("adminlevel")
rs.close
set rs=nothing

%>
<%
response.redirect("adminlogin.asp")
response.end
%>
