<%
IID=request("id")
dim name,pwd,email,person
dim sql
dim rs
set rs=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("mydb.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
sql = "DELETE FROM mytable WHERE id="+IID
rs.Open sql,conn,1,1
set rs=nothing
response.write "<script language=JavaScript>" & chr(13) & "alert('同学信息已经删除！');"&"window.location.href = 'index.asp'"&" </script>" 
%>
