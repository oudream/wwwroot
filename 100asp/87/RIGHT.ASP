<body  bgcolor=AntiqueWhite>
<%
dim IID,cate
if request("id")="" then
	cate=1
	cate="办公室"
else
	IID=request("id")
	cate=request("cate")
end if
session("IID")=IID
session("cate")=cate
%>
<table>
<tr align=center><td bgcolor=#ffcc00 width=80><%Response.write cate%></td>
<td colspan=5 width=520><a href="manage.asp?curraction=INS">添加</a>&nbsp;&nbsp;<a href="manage.asp?curraction=DEL">删除</a></td></tr>
<%
set rs=server.createobject("adodb.recordset") 
conn = "DBQ=" + server.mappath("message.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 
if IID="" then
	sql="select * from mytable where cateid=1"
else
	sql="select * from mytable where cateid="&IID
end if
rs.open sql,conn,1,1
if not rs.EOF then
%><tr align=center><td>姓名</td><td>性别</td><td>出生年月</td><td>学历</td><td>职务</td><td>部门</td></tr>
<%do while not rs.EOF
%>
<tr align=center>
<td><%=rs("name")%></td>
<td width=40><%=rs("sex")%></td>
<td width=140><%=rs("borndate")%></td>
<td width=140><%=rs("term")%></td>
<td width=140><%=rs("func")%></td>
<td width=60><%=rs("cate")%></td>
</tr>
<%
rs.MoveNext
loop
end if
set rs=nothing 
set conn=nothing
%></table>
</body>