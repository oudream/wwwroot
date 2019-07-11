<%
dim rs
dim sql
dim selected
dim conn
IID=request("IID")
selected=request.form("options")
set rs=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("message.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
sql="select * from problem where id="+IID
rs.open sql,conn,1,1
if rs("answer")=selected then
	if session("acceptnum")="" then
		session("acceptnum")=1
	else
		session("acceptnum")=session("acceptnum")+1
	end if
end if
set rs=nothing
set conn=nothing
Response.Redirect "index.asp"
%>