<%
  ' on error resume next
StrSQL= "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & Server.Mappath("../data/news3000.mdb")
'StrSQL="DSN=news;"
set conn=server.createobject("ADODB.CONNECTION")
conn.open StrSQL
%>