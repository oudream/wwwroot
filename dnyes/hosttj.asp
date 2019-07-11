<TABLE width=70% border=0 cellPadding=0 cellSpacing=0>
  <TBODY>
    <TR> 
      <TD height="8"><img src="1.gif" width="1" height="1"></TD>
    </TR>
  </TBODY>
</TABLE>
<table width="188" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td colspan="2"></td>
  </tr>
<%
sql="select top 8 * from subs where iftj=true order by id"
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
for i=1 to rs.recordcount%>
  <tr> 
    <td width="4" valign="top">&nbsp;</td>
    <td height="18" valign="top"> <a href="showsubs.asp?id=<%=rs("id")%>" title="<%=rs("subsname")%>" target="_blank"><font color="<%=rs("color")%>"><%=rs("subsname")%></font></a></td>
    <td align="right" width="80"><font color="<%=rs("color")%>"><%if rs("ifdisc")=true then%> <%response.write csng(rs("price"))*session("y_it_userdiscount")%> <%else%> <%response.write csng(rs("price"))%> <%end if%>
      /<%=rs("subsunit")%> </font></td>
  </tr>
<%rs.movenext
next
rs.close
set rs=nothing%>

</table>