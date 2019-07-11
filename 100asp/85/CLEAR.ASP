<%
set rs=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("mydb.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
sql="delete * from shop"
rs.open sql,conn,1,1
set rs=nothing
set conn=nothing
response.write "<script language=JavaScript>" & chr(13) & "alert('你的购物蓝已经清空！');" & "history.back()" & "</script>" 
%>