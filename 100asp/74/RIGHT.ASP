<body  bgcolor=AntiqueWhite>
<%
dim IID
IID=request("id")
set rs=server.createobject("adodb.recordset") 
conn = "DBQ=" + server.mappath("mybook.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 
if IID="" then
	sql="select * from mytable where id=1"
else
	sql="select * from mytable where id="&IID
end if
rs.open sql,conn,1,1
if not rs.EOF then
%>
<table width=600 align=left>
<tr><td bgcolor=#ffcc00 align=center><font size=+2><U><STRONG><%=rs("title")%></STRONG></U></font></td></tr>
<tr><td bgcolor=Linen><%=rs("content")%></td></tr></table>
<%
end if
set rs=nothing 
set conn=nothing
%>
</body>