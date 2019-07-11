<%
user=request("user")
email=request("email")
towho=request("towho")
content=request("content")
addtime=date()
set rs=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("message.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
sql="insert into mytable(user,email,towho,addtime,content) values('"&user&"','"&email&"','"&towho&"','"&addtime&"','"&content&"')"
rs.open sql,conn,1,1
set rs=nothing
set conn=nothing
response.write "<script language=JavaScript>" & chr(13) & "alert('¼ÇÂ¼ÒÑÌí¼Ó£¡');" & "history.back()" & "</script>" 
%>