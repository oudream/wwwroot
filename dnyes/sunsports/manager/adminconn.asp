<%
  ' on error resume next
if session("user_adminlevel")<>2 then 
response.Redirect("error.asp?err=001")
response.End()
end if
StrSQL= "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & Server.Mappath("../data/sports.mdb")
'StrSQL="DSN=news;"
set conn=server.createobject("ADODB.CONNECTION")
conn.open StrSQL
%>