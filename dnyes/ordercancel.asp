<!--#include file="mymefaq5411jkjkh/favorible/showme.asp" -->
<%
if session("y_it_uid")="" then
response.redirect "index.asp"
response.end
end if
%>
<%
id=request("id")
if id="" then
response.redirect "viewmyorders.asp"
response.end
end if
%>
<%
sql="select * from orders where id="&id
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs("username")<>session("y_it_uid") then
rs.close
set rs=nothing
response.redirect "viewmyorders.asp"
response.end
end if
rs.close
set rs=nothing
sql="delete from orders where id="&id
conn.Execute(sql)
conn.close
set conn=nothing
response.Redirect("viewmyorders.asp")
response.end
%>

