<TABLE width=70% border=0 cellPadding=0 cellSpacing=0>
  <TBODY>
    <TR> 
      <TD height="8"><img src="1.gif" width="1" height="1"></TD>
    </TR>
  </TBODY>
</TABLE>
<table width="200" border="0" cellspacing="0" cellpadding="0">
<%
sqlnews="select id,subject,message,color from message where newstype='±¾Õ¾¶¯Ì¬' order by id desc" 
Set rsnews=Server.CreateObject("ADODB.RecordSet") 
rsnews.Open sqlnews,conn,1,1 
n=0
do while not rsnews.eof
%>
                    <tr> 
                      
    <td width="12" height="20"><img src="images/dot-520.gif" width="12" height="20"> 
    </td>
                      <td valign="top">
					  <p style="line-height: 130%">
<a href="shownews.asp?id=<%=rsnews("id")%>" title="<%=rsnews("subject")%>" target="_blank"> 
<font color="<%=rsnews("color")%>">
<% if len(rsnews("subject"))>16 then%> <%=left(rsnews("subject"),15)%>..<%else%> 
<%=rsnews("subject")%> <%end if%></font></a></td>
                    </tr>
<%
n=n+1
rsnews.movenext
if n>=6 then exit do
loop
rsnews.close
set rsnews=nothing
%>
                  </table>
<TABLE width=70% border=0 cellPadding=0 cellSpacing=0>
  <TBODY>
    <TR> 
      <TD height="8"><img src="1.gif" width="1" height="1"></TD>
    </TR>
  </TBODY>
</TABLE>
