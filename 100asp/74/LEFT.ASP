<html>
<head>
<title>¶ÁÊיÊÒ</title>
</head>
<body>
<table border="0" align=right>
<tr><td align=center bgcolor=LightSteelBlue><font size=+1><STRONG>Í¼ÊיÄ¿Â¼</STRONG></font></td></tr>
<% 
set rs=server.createobject("adodb.recordset") 
conn = "DBQ=" + server.mappath("mybook.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 
sql="select * from mytable" 
rs.open sql,conn,1,1 
if not rs.EOF or rs.BOF then
do while not rs.EOF or rs.BOF
%>
<tr><td width=120 align=left bgcolor=MediumAquamarine>
<a href="right.asp?id=<%=rs("id")%>" target=main>
<%=rs("title")%></a>
</td></tr>
<%
rs.movenext
loop
end if
set rs=nothing 
set conn=nothing
%>
</TABLE> 