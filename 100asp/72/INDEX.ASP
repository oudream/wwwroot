<html>
<head>
<title>数据库的结构显示</title>
</head>
<body bgcolor=Peachpuff>
<p align=center><font face="华文行楷" size=+2><STRONG>sysobjects表结构的显示</STRONG></font><p>
<table border="1" align=center width=520 bgcolor=LightSteelBlue>
<% 
set rs=server.createobject("adodb.recordset") 
set rstemp=server.createobject("adodb.recordset") 
conn="provider=Sqloledb;server=username;uid=USERID;pwd=PASSWORD;database=pubs"
'显示sysobjects的表结构
sql="select id from sysobjects where name='sysindexes'" 
rs.open sql,conn,1,1
dim IID 
if rs.EOF then
	Response.Write "要显示的表不存在!"
else
	IID=rs("id")
SQLTEMP="SELECT * FROM [syscolumns] WHERE id ='"&IID&"' "
rstemp.Open sqltemp,conn,1,1
%>
<tr>
<td width=180>字段名称</td>
<td width=120>字段类型</td>
<td width=100>字段长度</td>
<td width=120>是否充许为空</td>
<%do while not rsTEMP.EOF%>
<TR>
<tr>
<td><%=rstemp("name")%></td>
<td><%
xtype=rstemp("xtype")
set rstemp1=server.createobject("adodb.recordset") 
sql="select name from [systypes] where xtype='"&xtype&"'"
rstemp1.Open sql,conn,1,1
Response.Write rstemp1("name")
rstemp1.Close
%></td>
<td><%=rstemp("length")%></td>
<td><%=rstemp("isnullable")%></td>
</tr>
<%
rsTEMP.MoveNext
loop
end if
rs.Close
rstemp.Close
set rstemp=nothing
set rs=nothing 
set conn=nothing
%>
</TABLE> 