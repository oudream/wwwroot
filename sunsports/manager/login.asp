<!--#include file="adminconn.asp" -->
<%
if request("option")="logout" then
            session("user_uid")=""
            session("user_pwd")=""
			session("user_name")=""
            session("user_adminlevel")=""
            session("user_email")=""
            session("user_tel")=""
            session("user_zip")=""
			session("user_pid")=""
			session("manager_pid")=""
			session("manager_tid")=""
			session("manager_tname")=""
url="../index.asp"
response.redirect url
response.end
end if
%>
<%
uid=trim(request("uid"))
pwd=trim(request("pwd"))
if uid="" or pwd="" then
%>
<SCRIPT LANGUAGE=JavaScript>
alert(' Please enter a valid Userid and Password . ');
window.history.go( -1 );
</SCRIPT>
<%
response.end
end if

sql="select * from player where uid='"&uid&"'"
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
if rs.eof or rs.bof then
%>
<SCRIPT LANGUAGE=JavaScript>
alert(' Please enter a valid Userid and Password . ');
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
rs.close
set rs=nothing

url=request("url")
response.redirect url
'response.write session("y_it_uid") & "<br>" & session("y_it_pwd") & "<br>" & session("y_it_userdiscount")
'response.end
%>
