<%
  ' on error resume next
StrSQL= "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & Server.Mappath("ffsd65adf5das6.asp")
'StrSQL="DSN=news;"
set conn=server.createobject("ADODB.CONNECTION")
conn.open StrSQL
%>