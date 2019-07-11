<%
dim rs
dim sql
dim selected
dim conn
selected=request.form("options")
set rs=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("search.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
sql="update search set sel"&selected&"=sel"&selected&"+1 where id=1"
rs.open sql,conn,1,1
set rs=nothing
set conn=nothing
response.write "<script language=JavaScript>" & chr(13) & "alert('您已经投了票，谢谢您的支持！');" & "history.back()" & "</script>" 
%>