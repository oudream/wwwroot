<%
Set Conn=Server.CreateObject("ADODB.Connection") 
Connstr="DBQ="+server.mappath("mymefaq5411jkjkh/asdewr2482f/safkk355f.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)}"
Conn.Open Connstr 
if session("y_it_userdiscount")="" then session("y_it_userdiscount")=1
%>
