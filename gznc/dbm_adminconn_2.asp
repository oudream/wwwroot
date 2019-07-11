<%
  ' on error resume next
if not ( session("user_adminlevel")=1 or session("user_adminlevel")=2 or session("user_adminlevel")=9 ) then 
response.Redirect("dbm_adminlogin.asp")
response.End()
end if
StrSQL= "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & Server.Mappath("showout/data/dbmdata.mdb")
'StrSQL="DSN=news;"
set conn=server.createobject("ADODB.CONNECTION")
conn.open StrSQL
%>