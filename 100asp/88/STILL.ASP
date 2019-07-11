<%
dim name,author
if request("text5")="" then
	response.write "<script language=JavaScript>" & chr(13) & "alert('请输入图书编号！');" & "history.back()" & "</script>" 
	Response.End
end if
id=request("text5")
dim sql
dim rs
dim conn
dim islend
islend=false
set rs=server.createobject("adodb.recordset") 
conn = "DBQ=" + server.mappath("book.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 
sql="update mytable set islend="&islend&",username='' where id="+id
rs.Open sql,conn,1,1
response.write "<script language=JavaScript>" & chr(13) & "alert('图书归还成功！');" & "history.back()" & "</script>" 
%>
