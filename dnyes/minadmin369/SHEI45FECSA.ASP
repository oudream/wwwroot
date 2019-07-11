<%
if session("y_eshop_adminuser")="" then 
response.redirect "../error.asp?error=011"
response.end
else
Set Conn=Server.CreateObject("ADODB.Connection") 
Connstr="DBQ="+server.mappath("../mymefaq5411jkjkh/asdewr2482f/safkk355f.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)}"
Conn.Open connstr 
end if
%>
