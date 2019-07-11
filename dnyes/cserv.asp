<TABLE width=70% border=0 cellPadding=0 cellSpacing=0>
  <TBODY>
    <TR> 
      <TD height="8"><img src="1.gif" width="1" height="1"></TD>
    </TR>
  </TBODY>
</TABLE>
<table width="212" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td colspan="2"></td>
  </tr>
  <%
sqlnews="select * from message where newstype='客服中心' order by id desc" 
Set rsnews=Server.CreateObject("ADODB.RecordSet") 
rsnews.Open sqlnews,conn,1,1 
n=0
do while not rsnews.eof
%>
  <tr> 
    <td width="20" height="18" valign="top"> <div align="center"></div></td>
    <td valign="top"><a href="shownews.asp?id=<%=rsnews("id")%>" title="<%=rsnews("subject")%>" target="_blank"> 
     <font color="<%=rsnews("color")%>">
	  <% if len(rsnews("subject"))>20 then%>
      <%=left(rsnews("subject"),25)%>..
      <%else%>
      <%=rsnews("subject")%> 
      <%end if%></font>
      </a></td>
  </tr>
  <%
n=n+1
rsnews.movenext
if n>=8 then exit do
loop
rsnews.close
set rsnews=nothing
%>
  <tr> 
    <td colspan="2"></td>
  </tr>
</table>
