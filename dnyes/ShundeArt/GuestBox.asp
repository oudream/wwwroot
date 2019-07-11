<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<html>
<head>
<title></title>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%=request.cookies(Forcast_SN)("content")%>
<%
Dim action,newsID,rs,Content
dim sql
dim conn
Action=LCase(Request.QueryString("Action"))
reviewid=Request.QueryString("reviewid")

if not IsNumeric(reviewid)then
	response.write "<script>alert('非法参数');history.back()</script>"
	response.end
end if

If request("action")="modify" Then
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_Review_Table &" where reviewid="&reviewid
	rs.open sql,conn,1,1
	If Not rs.Eof Then
		Content=rs("Content")
	End If
	Response.Write Content
End If
%>
</body>
</html>