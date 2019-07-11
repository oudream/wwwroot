<table width="171" height="250" border="0" cellpadding="0" cellspacing="0">
              <tr valign="top"> 
                <td colspan="2">租用托管</td>
              </tr>
<%
sqlnews="select * from myfaq where newstype='租用托管' order by id desc" 
Set rsnews=Server.CreateObject("ADODB.RecordSet") 
rsnews.Open sqlnews,conn,1,1 
n=0
do while not rsnews.eof
%> 
              <tr> 
                <td width="32" height="18" valign="top"> 
                  <div align="center"><img src="image/icon.gif" border="0"></div>
                </td>
                <td width="157" valign="top"><a href="showfaq.asp?id=<%=rsnews("id")%>" title="<%=rsnews("subject")%>" target="_blank"> 
                  <font color=<%=rsnews("color")%>><% if len(rsnews("subject"))>13 then%> <%=left(rsnews("subject"),11)%>..<%else%> 
                  <%=rsnews("subject")%> <%end if%></font></a></td>
              </tr>
<%
n=n+1
rsnews.movenext
if n>=10 then exit do
loop
rsnews.close
set rsnews=nothing
%>
              <tr> 
                <td>&nbsp;</td>
                <td><div align="right"><a href="viewfaq.asp?newstype=租用托管" >更多&gt;&gt;</a></div></td>
              </tr>
            </table>