<!--#include file="dbm_wellcome_conn.asp" -->
<%
if request("option")="logout" then
            session("guest_id")=0
			session("guest_gid")=""
			session("guest_name")=""
            session("guest_adminlevel")=0
			session("user_allname")=""
response.Write("<SCRIPT LANGUAGE=JavaScript> window.close();</SCRIPT>")
response.end
end if
%>
<%
gid=trim(request("gid"))
pwd=trim(request("pwd"))

sql="select * from dbm_guest where gid='"&gid&"'"
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
if rs.eof or rs.bof then
%>
<SCRIPT LANGUAGE=JavaScript>
alert(' «ÎÃÓ–¥”√ªß’À∫≈! ');
window.history.go( -1 );
</SCRIPT>
<%
response.end
end if

if pwd<>rs("pwd") then
%>
<SCRIPT LANGUAGE=JavaScript>
alert(' Please enter a valid guestid and Password . ');
window.history.go( -1 );
</SCRIPT>
<%
response.end
end if

session("guest_id")=rs("guest_id")
session("guest_gid")=rs("gid")
session("guest_name")=rs("allname")
session("user_adminlevel")=9
rs.close
set rs=nothing

%>
<%
response.redirect("index.asp")
response.end
%>
