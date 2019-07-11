<%@ Language=VBScript %>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->

<%
set rs=server.CreateObject("adodb.recordset")
rs.Source="select top "& NewsNum &" * from "& db_News_Table &" order by NewsID DESC"
rs.Open rs.Source,conn,1,1

set fs=server.CreateObject("Scripting.FileSystemObject")
FilePath=server.MapPath("LastNewsXP.asp")
set JSFile=fs.CreateTextFile(FilePath)
JSFile.writeline("<table cellpadding=3 cellspacing=0 border=0 >")

while not rs.EOF
	JSFile.writeline("	<tr>")
	JSFile.writeline("		<td width=100% >")
	StrNews="			<a href='ReadNews.asp?NewsID="& rs("NewsID") &"' target=_blank>"& trim(rs("Title")) &"</a>("& year(rs("UpdateTime")) &"-"& Month(rs("UpdateTime")) &"-"& Day(rs("UpdateTime")) &")"
	JSFile.writeline StrNews
	JSFile.writeline("		<td width=100% >")
	JSFile.writeline("	</tr>")
	rs.MoveNext
wend

JSFile.writeline("</table>")

rs.Close
set rs=nothing
conn.close
set conn=nothing

set JSFile=nothing
set fs=nothing
%>

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="refresh" content=""&freetime&";url=addnews0.asp">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
</head>
<body>
<p>　</p>
<p>　</p>
<p>　</p>
<div align="center">
	<center>
		<table border="1" width="450" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="<%=border%>">
			<tr>
				<td colspan=2 align="center" bgcolor="<%=m_top%>">
					<p></p>
					<p>首页调用ASP已经更新！</p>
					<p></p>
				</td>
			</tr>
		</table>
	</center>
</div>
</body>
</html>