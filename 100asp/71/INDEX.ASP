<html>
<head>
<title>数据库的结构显示</title>
</head>
<body bgcolor=Peachpuff>
<p align=center><font face="华文行楷" size=+2><STRONG>实例&nbsp;&nbsp;数据库结构的显示</STRONG></font><p>
<table border="1" align=center width=100 bgcolor=LightSteelBlue><tr>
<% 
set rs=server.createobject("adodb.recordset") 
'对于SQL中的系统数据库PUBS
conn="provider=Sqloledb;server=servername;uid=username;pwd=password;database=pubs"
'servername：所用的服务器名
'UserID：用户名，常用为SA
'password：数据库密码，一般情况下为空
'显示数据库中的所有用户表
'sql="select * from sysobjects where xtype='U'"
'显示数据库中PUBS的所有系统表 
sql="select * from sysobjects where xtype='S'" 
'对于ACCESS中的数据库凡MYDB
'conn = "DBQ=" + server.mappath("message.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 
'sql="select * from MsysObjects where type=1" 
rs.open sql,conn,1,1
howmanyfields=rs.Fields.Count-1
for i=0 to howmanyfields
%>
<TD>
<STRONG><%=rs(i).name%></STRONG>
</td>
<%
next
%>
</tr>
<%do while not rs.EOF%>
<TR>
<%
for i=0 to howmanyfields
	thisvalue=rs(i)
	if isnull(thisvalue) then
		thisvalue="&nbsp;"
	end if
	%>
	<td><%=thisvalue%></td>
	<%
next
%>
</tr>
<%
rs.MoveNext
loop
rs.Close
set rs=nothing 
set conn=nothing
%>
</TABLE> 