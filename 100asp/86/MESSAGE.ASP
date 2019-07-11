<%
if session("name")="" then
	Response.Write "您还没有注册，请先注册!"
	Response.End
end if
%>
<body>
<p align=center>您购买的商品有：</p>
<table width=450 align=center>
<%
dim curid
set rs=server.createobject("adodb.recordset") 
conn = "DBQ=" + server.mappath("mydb.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 
sql="select * from mytable" 
rs.open sql,conn,1,1
curid=0
if not rs.EOF or rs.BOF then 
%>
<tr bgcolor=LawnGreen><td WIDTH=45%>手机型号</td><td WIDTH=25%>生产厂家</td><td WIDTH=15%>价格</td><td WIDTH=15%>数量</td></tr> 
<%
do while not rs.eof
curid=curid+1
if isnumeric(request(curid))<>0 then
%>
	<tr bgcolor=LightGoldenrodYellow><td><%=rs("name")%></td> 
	<td><%=rs("publish")%></td> 
	<td><%=rs("price")%>元</td> 
	<td><%if request(curid)<>"" then 
	Response.Write request(curid)
	else 
	Response.write "0"
	end if	%>件
	</td></tr>
<% 	
	end if
	rs.movenext
	loop 
end if 
set rs=nothing 
set conn=nothing 
%></table>
