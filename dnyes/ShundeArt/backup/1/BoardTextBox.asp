<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html>
<head>
<title></title>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"><LINK href=site.css rel=stylesheet>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" class="black9">
<%
Dim action,ID,rs,Content
dim sql
dim conn
Action=LCase(Request.QueryString("Action"))
ID=Request.QueryString("ID")

If request("action")="modify" Then
set rs=server.createobject("adodb.recordset")
sql="select * from "& db_board_Table &" where id="&id
rs.open sql,conn,1,1
	If Not rs.Eof Then
		Content=rs("Content")
	End If
	Response.Write Content
End If
%>
</body>
</html>