<table width="167" height="243" border="0" cellpadding="0" cellspacing="0">
              <tr valign="top"> 
                <td colspan="2">网站建设</td>
              </tr>
<%
sqlnews="select * from myfaq where newstype='网站建设' order by id desc" 
Set rsnews=Server.CreateObject("ADODB.RecordSet") 
rsnews.Open sqlnews,conn,1,1 
n=0
do while not rsnews.eof
%> 
              <tr> 
                <td width="20" height="18" valign="top"> 
                  <div align="center"><img src="image/icon.gif" border="0"></div>
                </td>
                <td valign="top"><a href="showfaq.asp?id=<%=rsnews("id")%>" title="<%=rsnews("subject")%>" target="_blank"> 
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
                <td><div align="right"><a href="viewfaq.asp?newstype=网站建设" >更多&gt;&gt;</a></div></td>
              </tr>
            </table>