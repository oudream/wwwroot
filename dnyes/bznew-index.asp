<TABLE width=70% border=0 cellPadding=0 cellSpacing=0>
  <TBODY>
    <TR> 
      <TD height="8"><img src="1.gif" width="1" height="1"></TD>
    </TR>
  </TBODY>
</TABLE>
<table width="98%" border="0" cellpadding="0" cellspacing="0">
  <%
sqlnews="select id,subject,message,color from message where newstype='本站动态' order by id desc" 
Set rsnews=Server.CreateObject("ADODB.RecordSet") 
rsnews.Open sqlnews,conn,1,1 
n=0
do while not rsnews.eof
%>
  <tr> 
    <td width="4" align="right" valign="top">&nbsp;</td>
    <td height="19"><a href="shownews.asp?id=<%=rsnews("id")%>" title="<%=rsnews("subject")%>" target="_blank"> 
      <font color=<%=rsnews("color")%>> 
      <% if len(rsnews("subject"))>15 then%>
      <%=left(rsnews("subject"),14)%>.. 
      <%else%>
      <%=rsnews("subject")%> 
      <%end if%>
      </font></a></td>
  </tr>
  <%
n=n+1
rsnews.movenext
if n>=6 then exit do
loop
rsnews.close
set rsnews=nothing
%>
  <tr> 
    <td colspan="2"><div align="right"><a href="viewnews.asp?newstype=本站动态" >更多&gt;&gt;</a></div></td>
  </tr>
</table>
<TABLE width=70% border=0 cellPadding=0 cellSpacing=0>
  <TBODY>
    <TR> 
      <TD height="7"><img src="1.gif" width="1" height="1"></TD>
    </TR>
  </TBODY>
</TABLE>
