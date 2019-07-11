<%@ Language=VBScript %>
<body  bgcolor=AntiqueWhite><%
set rs=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("message.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
if request("id")="" then
	response.write "<script language=JavaScript>" & chr(13) & "alert('非法操作！');" & "history.back()" & "</script>" 
else
	id=request("id")
	sql="delete from mytable where id="+id
	rs.Open sql,conn,1,1
	Response.Write "记录已经删除!"
end if
set rs=nothing
set conn=nothing
%>
</body>
