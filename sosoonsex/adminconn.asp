<%
  ' on error resume next
if not ( session("user_adminlevel")=1 or session("user_adminlevel")=3 ) then 
response.Redirect("adminlogin.asp")
response.End()
end if
StrSQL= "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & Server.Mappath("ffsd65adf5das6.asp")
'StrSQL="DSN=news;"
set conn=server.createobject("ADODB.CONNECTION")
conn.open StrSQL
%>