<%
dim name,author
if request("text1")="" then
	response.write "<script language=JavaScript>" & chr(13) & "alert('请输入书名！');" & "history.back()" & "</script>" 
	Response.End
end if
if request("text2")="" then
	response.write "<script language=JavaScript>" & chr(13) & "alert('请输入作者！');" & "history.back()" & "</script>" 
	Response.End
end if
name=request("text1")
author=request("text2")
dim sql
dim rs
dim conn
set rs=server.createobject("adodb.recordset") 
conn = "DBQ=" + server.mappath("book.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 
sql="insert into mytable(name,author) values('"&name&"','"&author&"')"
rs.Open sql,conn,1,1
Response.Write "记录添加成功!"
%>