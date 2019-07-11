<%
  ' on error resume next
StrSQL= "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & Server.Mappath("database/sports.mdb")
'StrSQL="DSN=news;"
set conn=server.createobject("ADODB.CONNECTION")
conn.open StrSQL
%>