<%
dim name,author
if request("text3")="" then
	response.write "<script language=JavaScript>" & chr(13) & "alert('������ͼ���ţ�');" & "history.back()" & "</script>" 
	Response.End
end if
id=request("text3")
dim sql
dim rs
dim conn
set rs=server.createobject("adodb.recordset") 
conn = "DBQ=" + server.mappath("book.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 
sql="delete  from mytable where id="+id
rs.Open sql,conn,1,1
response.write "<script language=JavaScript>" & chr(13) & "alert('ͼ��ɾ���ɹ���');" & "history.back()" & "</script>" 
%>