<!--#include file="Top_sun.asp"-->
<%
eid=request("eid")
if eid="" then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Please select the Events you want to View at first ! ');</SCRIPT>")
response.End()
end if
%>
<%
sql="select * from events where eid="&eid&" order by eid desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 
if rs.eof or rs.bof then response.Redirect("error.asp?err=003")
%>
<table border="0" cellspacing="0" cellpadding="0" width="776" align="center">
<%
rs.close
set rs=nothing
%>
<%
sql="select * from events where eid<>"&eid&" order by eid desc" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if rs.eof or rs.bof then response.Redirect("error.asp?err=003")

if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=10
rs.AbsolutePage=pagecount
%>
<%
do while not rs.eof 
%>
  <tr> 
    <td colspan="4"><a href="events_show.asp?eid=<%=rs("eid")%>"><%=rs("name")%></a></td>
  </tr>
<%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
rs.movefirst
%>
</table>
<%
rs.close
set rs=nothing
%>
<!--#include file="Copyright.asp"-->