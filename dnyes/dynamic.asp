<table width="170" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td colspan="3">
          </td>
        </tr>
        <tr> 
          <td colspan="3"><img src="image/hyxx.gif" width="170"></td>
        </tr>
        <tr> 
          <td width=1 bgcolor="#285278"></td>
          <td>
            <table width="170" border="0" cellspacing="0" cellpadding="0" bgcolor="#D2E8FF" style="BORDER-RIGHT: #0079ce 1px solid; BORDER-LEFT: #0079ce 1px solid">              <tr> 
                <td colspan="2"></td>
              </tr>
<%
sqlnews="select id,subject,message from message where newstype='行业信息' order by id desc" 
Set rsnews=Server.CreateObject("ADODB.RecordSet") 
rsnews.Open sqlnews,conn,1,1 
n=0
do while not rsnews.eof
%> 
              <tr> 
                <td width="20" height="18" valign="top"> 
                  <div align="center"><img src="image/icon.gif" border="0"></div>
                <td valign="top"><a href="shownews.asp?id=<%=rsnews("id")%>" title="<%=rsnews("subject")%>" target="_blank"> 
                  <% if len(rsnews("subject"))>13 then%> <%=left(rsnews("subject"),11)%>..<%else%> 
                  <%=rsnews("subject")%> <%end if%></a></td>
              </tr>
<%
n=n+1
rsnews.movenext
if n>=7 then exit do
loop
rsnews.close
set rsnews=nothing
%>
              <tr> 
                <td colspan="2"></td>
              </tr>
            </table>
          </td>
          <td width=1 bgcolor="#285278"></td>
        </tr>
        <tr> 
          <td colspan="3"><img src="image/d_1.gif" width="170" height="10"></td>
        </tr>
      </table><br>
      