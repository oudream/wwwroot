<%
dim IID,upNUM
IID=request("select1")
upNUM=request("select2")
dim sql
dim rs,rssel
dim conn
dim nowday
nowday=date()
set rssel=server.CreateObject("adodb.recordset")
set rs=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("message.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 
sql="select nowday from mytable where id="+IID
rssel.Open sql,conn,1,1
if rssel("nowday")=date() then
	rssel.Close
	set rssel=nothing
	Response.Write "<script language=JavaScript>" & chr(13) & "alert('这名员工今日已登记！');" & "history.back()" & "</script>" 
else 
	rssel.Close
	set rssel=nothing
	sql="update mytable set "&upNUM&"="&upNUM&"+1, nowday='"&nowday&"' where id="+IID
	rs.Open sql,conn,1,1
	response.write "<script language=JavaScript>" & chr(13) & "alert('员工登记成功！');" & "window.location.href = 'index.asp'"&" & "</script>" 
end if
%>
