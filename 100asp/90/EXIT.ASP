<%
if session("name")="" then
	Response.Write "<script language=JavaScript>" & chr(13) & "alert('非法用户！');" & "history.back()" & "</script>" 
	Response.End
else
	if session("name")="lyf" then
		Response.Write "<script language=JavaScript>" & chr(13) & "alert('感谢您的光临，欢迎下次再来！！');" &"window.close()"&"</script>" 
	else
		dim sql
		dim rs
		name=session("name")
		set rs=server.createobject("adodb.recordset")
		conn = "DBQ=" + server.mappath("mydb.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
		sql="delete * from activetable where name='"&name&"'"
		Response.Write sql
		rs.Open sql,conn,1,1
		set rs=nothing
		session("name")=""
		Response.Write "<script language=JavaScript>" & chr(13) & "alert('感谢您的光临，欢迎下次再来！！');" & "window.close();" & "</script>" 
	end if
end if
%>