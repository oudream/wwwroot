<%
if session("name")="" then
	Response.Write "<script language=JavaScript>" & chr(13) & "alert('�Ƿ��û����������ύ��Ϣ��');" & "history.back()" & "</script>" 
	Response.End
else
dim message
dim sql
dim rs
if request("text1")="" then
	response.write "<script language=JavaScript>" & chr(13) & "alert('������������Ϣ��');" & "history.back()" & "</script>" 
	Response.End
end if
name=session("name")
message=request("text1")
set rs=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("mydb.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
sql="insert into message(name,message) values('"&name&"','"&message&"')"
Response.Write sql
rs.Open sql,conn,1,1
set rs=nothing
Response.Redirect "send.asp"
end if
%>
