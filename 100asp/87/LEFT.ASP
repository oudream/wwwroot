<html>
<head>
<title>人事管理</title>
</head>
<body bgcolor=Peachpuff>
<p align=center><font size=+1><STRONG>人事管理</STRONG></font><p>
<table border="0" align=center width=100 bgcolor=LightSteelBlue>

<% 
set rs=server.createobject("adodb.recordset") 
conn = "DBQ=" + server.mappath("message.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 
sql="select * from cate" 
rs.open sql,conn,1,1 
if rs.EOF or rs.BOF then
	Response.Write "<tr><td>数据库暂无信息</td></tr>"
else
do while not rs.EOF or rs.BOF
%>
<tr><td align=center bgcolor=MediumAquamarine>
<a href="right.asp?id=<%=rs("id")%>&cate=<%=rs("cate")%>" target=main>
<%=rs("cate")%></a>
</td></tr>
<%
rs.movenext
loop
end if
set rs=nothing 
set conn=nothing
%>
</TABLE> 