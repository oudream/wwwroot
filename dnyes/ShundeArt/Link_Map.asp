<table border=0 cellpadding=0 cellspacing=3 width="88" align="center"><%
set rs10=server.CreateObject("ADODB.RecordSet") 
rs10.Source="select top "& linkshownum &" * from "& db_link_Table &" where linktype=1 and pass=1 order by ID DESC "
rs10.Open rs10.Source,conn,1,1
if not rs10.EOF then
while not rs10.EOF
%>
  <tr> 
    <td height=31 align="center"><a href="<%=rs10("weburl")%>" target="_blank" title="<%=rs10("content")%>&#13;&#10;站长：<%=rs10("webmaster")%>&#13;&#10;申请时间：<%=rs10("dateandtime")%>"><img src="<%=rs10("logo")%>" width="88" height="31" border="0"></a></td>
  </tr>
  <%
rs10.MoveNext
wend
%>
<%
else
Response.Write "目前还没有"
end if
rs10.Close
set rs10=nothing
%>

 
</table>
</marquee>