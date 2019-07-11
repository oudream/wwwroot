<%
dim id,addtime
ID=request("checkid")
addtime=date()
dim rstemp,rs
dim name,price
set rstemp=server.createobject("adodb.recordset")
set rs=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("mydb.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
sqltemp="select name,price from mytable where id="+ ID
rstemp.Open sqltemp,conn,1,1
if rstemp.EOF or rstemp.BOF then
	response.write "<script language=JavaScript>" & chr(13) & "alert('这本图书已售完！');" & "history.back()" & "</script>" 
else
	name=rstemp("name")	
	price=rstemp("price")
end if
sql="insert into shop(name,price,addtime) values('"&name&"','"&price&"','"&addtime&"')"
rs.open sql,conn,1,1
response.write "<script language=JavaScript>" & chr(13) & "alert('这本图书已添加到你的购物蓝中！');" & "history.back()" & "</script>" 
set rs=nothing
set rstemp=nothing
set conn=nothing
%>
