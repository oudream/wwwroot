<%
dim name,author
if request("text4")="" then
	response.write "<script language=JavaScript>" & chr(13) & "alert('请输入图书编号！');" & "history.back()" & "</script>" 
	Response.End
end if
if request("text6")="" then
	response.write "<script language=JavaScript>" & chr(13) & "alert('请输入借阅人姓名！');" & "history.back()" & "</script>" 
	Response.End
end if
id=request("text4")
username=request("text6")
dim sql
dim rs
dim conn
dim islend
islend=true
set rs=server.createobject("adodb.recordset") 
conn = "DBQ=" + server.mappath("book.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 
sql="update mytable set islend="&islend&",username='"&username&"' where id="+id
'Response.Write sql
rs.Open sql,conn,1,1
response.write "<script language=JavaScript>" & chr(13) & "alert('图书借阅成功！');" & "history.back()" & "</script>" 
%>
