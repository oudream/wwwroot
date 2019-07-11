<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
<tr style="table-layout:fixed; word-break:break-all">
<%
set rs=server.CreateObject("ADODB.RecordSet") 
rs.Source="select * from "& db_Board_Table &" where inuse=1"
rs.Open rs.Source,conn,1,1
if rs.eof then %>
<div align="center">没有公告</div>');
<% else%>
<td width=100% align=center height=2><%=trim(htmlencode4(rs("title")))%><br><%=rs("dateandtime")%></td></tr>
<tr><td width=100% height=2><%=trim(CheckStr(rs("content")))%></td></tr>
<%
end if
Rs.Close
set Rs=nothing
%>
</table>