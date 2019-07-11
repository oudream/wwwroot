<TABLE width=70% border=0 cellPadding=0 cellSpacing=0>
  <TBODY>
    <TR> 
      <TD height="8"><img src="1.gif" width="1" height="1"></TD>
    </TR>
  </TBODY>
</TABLE>
<table width="212" border="0" cellpadding="0" cellspacing="0">
  <%
sqlnews="select * from sample_area order by id desc" 
Set rsnews=Server.CreateObject("ADODB.RecordSet") 
rsnews.Open sqlnews,conn,1,1 
n=0
do while not rsnews.eof
%>
  <tr> 
    <td width="20" height="16" align="right" valign="top"><img src="images/dot-520.gif" width="12" height="20"></td>
    <td><a href="http://<%=rsnews("addr")%>" title="<%=rsnews("sample_area_name")%>" target="_blank"> 
      <%=rsnews("sample_area_name")%> </a></td>
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
    <td colspan="2"><div align="right"><a href="Virtual-Hosting-Service-Sample.asp" ><font color="#FF0000">更多空间演示&gt;&gt;</font></a></div></td>
  </tr>
</table>
<TABLE width=70% border=0 cellPadding=0 cellSpacing=0>
  <TBODY>
    <TR> 
      <TD height="7"><img src="1.gif" width="1" height="1"></TD>
    </TR>
  </TBODY>
</TABLE>
