<table width="98%" border="0" cellpadding="0" cellspacing="0">
<%
sql="select * from tournament where type='cup' order by cid desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if rs.bof or rs.eof then response.Redirect("error.asp?err=1")
do while not rs.eof
%>
  <tr> 
    <td height="19"><a href="League.asp?cid=<%=rs("cid")%>" title="<%=rs("name")%>" target="_blank"> 
      <%=rs("name")%></a></td>
  </tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
  <tr> 
    <td>&nbsp;</td>
  </tr>
</table>