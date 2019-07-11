<%@ Language=VBScript %>
<!--#include file="database.asp"-->
<%
const MaxPerPage=15
dim curaction, IID, modify
curaction = trim(request("action"))
IID = request("id")
modify = false
if curaction = "edit" then
	modify = true
end if
dim sql
if curaction = "delete" then		'删除记录
	sql = "DELETE FROM mytable WHERE id=" + cstr(IID)
	conn.execute sql
	response.write "记录已经删除。"
end if
if curaction = "new" then		'添加记录
	if trim(request("txtname")) = "" then
		Response.Write " 请输入名称！"
	else
		sql = "INSERT INTO mytable(name,email,content) VALUES('"
		sql = sql + request("txtname") + "','"
		sql = sql + request("txtemail") + "','"
		sql = sql + request("txtcontent")+"')"
		response.write "记录已经添加到数据库。"
	end if
end if
if curaction = "edit" then		'修改记录
	if trim(request("txtname")) = "" then
		Response.Write " 请输入名称！"
	else
		sql = "UPDATE mytable SET "
		sql = sql + "name='" + htmlencode(request("txtname")) + "',"
		sql = sql + "email='" + htmlencode(request("txtemail")) + "',"
		sql = sql + "content='" + htmlencode(request("txtcontent")) + "'"
		sql = sql + " WHERE id = " + cstr(IID)
		response.write sql	
		response.write "记录已经更新到数据库。"
	end if
end if
%>
<html>
<head>
	<title>实例　通讯簿</title>
<script language="javascript">
function confirmDel(id){
if ( confirm("你确定要删除吗？")) {
window.location.href = "index.asp?action=delete&id=" + id
  }
}
</script>
</head>
<body background="bg.gif">
<p align=center><b><u><font size=+1>实例  通讯簿</u></p>
<% 
sqlsearch=""
name=request("name")
if name<>"" then
   sqlsearch=sqlsesarch & " and name like '%" & name & "%'"
end if   
set rs = server.createobject("adodb.recordset")
sql = "SELECT * FROM mytable where id<>0" + sqlsearch + " ORDER BY id DESC"
rs.open sql, conn, 1, 1
if rs.eof and rs.bof then
   response.write "没有记录"
else
 %>
<table width="85%" border="1" cellspacing="0" align="CENTER" bordercolor=<%= TBCOLOR %>>
<tr bgcolor=<%= THCOLOR %>>
<td><b>姓名</b></td>
<td><b>email</b></td>
<td><b>操作</b></td></tr>
<% 
dim curpage, i
if request("page")="" then
  curpage = 1
else
curpage = cint(request("page"))
end if
rs.pagesize = MaxPerPage
totalPut=rs.recordcount
rs.absolutepage = curpage
for i = 1 to rs.pagesize
%>
<tr><td><%= rs("name") %></td>
<td><%= rs("email") %></td>
<td><a href='index.asp?action=edit&id=<%= cstr(rs("id")) %>'>编辑</a> 
<a href='javascript:confirmDel(<%= rs("id") %>)'>删除</a></td></tr>
<% rs.movenext
if rs.eof then
i = i + 1
exit for
end if
next %>
</table>
<form method="POST" action="index.asp" id=form1 name=form1>
<% 
if (totalPut mod MaxPerPage)=0 then 
	n= totalPut \ MaxPerPage 
else 
	n= totalPut \ MaxPerPage + 1 
end if
Response.Write "共"+Cstr(totalPut)+"个好友　　"
for k=1 to n
if k=curpage then  
response.write "当前为：第[<b>"+Cstr(k)+"</b>]页 "
else
response.write "[<b>"+"<a href='index.asp'></a></b>] "  
end if  
next
%>
<br>姓名<input type="text" name="name" size="11">
<input type="text" name="xl" size="7"><input type="submit" value="提交" name="B1">
</form>
<%rs.close
end if
 %>
<hr>
<% 
if modify then
sql = "SELECT * FROM mytable WHERE id =" + cstr(IID)  
rs.open sql, conn, 1, 1
end if
%>
<form action="index.asp" method="POST" id=form2 name=form2>
<p align=center><b><% If modify then %>编辑：<% Else  %>添加：<% End If %></b></p>
<table align=center width="70%">
<tr>
<td><b>姓名：</b></td>   
<td><input name="txtname" size="15" value='<% If modify then   
response.write rs("name")   
end if   
%>'></td>   
</tr>   
<tr>   
<td><b>email：</b></td>   
<td><input name="txtemail" size="10" value='<% If modify then   
response.write rs("email")   
end if   
%>'></td>   
</tr>   
<tr> 
<td><b>备注：</b></td> 
<td><textarea name="txtcontent" wrap="off" cols="40" rows="6"><% If modify then
response.write rs("content") 
end if
 %></textarea></td> 
</tr> 
<tr><td align=center colspan="2">
<input type="Hidden" name="action" value='<% If modify then%>edit<% Else  %>new<% End If %>'>
<% If modify then %>
<input type="Hidden" name="id" value='<%= cstr(IID) %>'>
<% End If %>
<input type="Submit" name="save" value="保存">
</td></tr>
</table>
</font></b>
</form>
<% 
set rs = nothing
call endconnection()
 %>
</body>
</html>
