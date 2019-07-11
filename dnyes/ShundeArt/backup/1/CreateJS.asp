<%@ Language=VBScript %>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="chkuser.asp"-->
<%
set rs=server.CreateObject("adodb.recordset")
rs.Source="select top " & NewsNum &" * from "& db_News_Table &" order by NewsID DESC"
rs.Open rs.Source,conn,1,1

set fs=server.CreateObject("Scripting.FileSystemObject")
FilePath=server.MapPath("LastNewsXP.js")
set JSFile=fs.CreateTextFile(FilePath)
JSFile.writeline("document.write(""<table cellpadding=3 cellspacing=0 border=0 >"")")

while not rs.EOF
	JSFile.writeline("document.write(""   <tr>"")")
	JSFile.writeline("document.write(""     <td width=100% >"")")
	StrNews=" <a href='http://"& xpurl &"/ReadNews.asp?NewsID=" & rs("NewsID") &  "' target=_blank>" & trim(rs("Title")) & "</a>(" & year(rs("UpdateTime")) & "-" & Month(rs("UpdateTime")) &"-"& Day(rs("UpdateTime")) & ")"
	JSFile.writeline("document.write(""" &  StrNews & """)")
	JSFile.writeline("document.write(""     <td width=100% >"")")
	JSFile.writeline("document.write(""   </tr>"")")
	rs.MoveNext
wend

JSFile.writeline("document.write(""</table>"")")
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
<meta http-equiv="refresh" content=""&freetime&";url=createasp.asp">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
</head>
<body>
<p></p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<table border="1" width="450" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="<%=border%>" align=center>
  <tr> 
    <td colspan=2 align="center" bgcolor="<%=m_top%>"> 
      <p>&nbsp;</p>
      <p>已经生成 LastNewsXP.js，请在首页用调用该文件！</p>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>
</body>
</html>