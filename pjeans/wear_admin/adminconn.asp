<%
  ' on error resume next
if session("user_adminlevel")<>1 then 
response.Redirect("adminlogin.asp")
response.End()
end if
StrSQL= "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & Server.Mappath("../database/wearbysosoon.mdb")
'StrSQL="DSN=news;"
set conn=server.createobject("ADODB.CONNECTION")
conn.open StrSQL
%>