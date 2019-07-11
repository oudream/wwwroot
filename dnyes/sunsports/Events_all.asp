<!--#include file="Top_sun.asp"-->
<table width="776" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="181">&nbsp;</td>
    <td> <%
sql="select * from events order by eid desc" 
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
%> <table border="0" cellspacing="0" cellpadding="0" width="100%" align="center">
        <tr align="center"> 
          <td height="22" colspan="6" bgcolor="#FFFFFF"><strong><font color="#000000" size="3" face="Verdana, Arial, Helvetica, sans-serif">NEWS 
            TODAY </font></strong></td>
        </tr>
        <tr> 
          <td colspan="4"> <table width="100%" border="0" cellspacing="2" cellpadding="2">
              <tr> 
                <td height="22" align="center" valign="middle"><a href="events_show.asp?eid=<%=rs("eid")%>"><%=rs("name")%></a></td>
              </tr>
              <tr> 
                <td height="22" align="center" valign="middle"> <%
if trim(rs("photo"))<>"" then
%> <a href="events_show.asp?eid=<%=rs("eid")%>"><img src=<%=rs("photo")%> border="0" width="300" height="120"></a> 
                  <%
end if
%> &nbsp;</td>
              </tr>
              <tr align="left" valign="top"> 
                <td height="100">NOTE : 
                  <%
		  if len(rs("memo"))>100 then
		  response.Write(left(rs("memo"),98)&"...")
		  else
		  response.Write(rs("memo"))
		  end if
		  %> </td>
              </tr>
            </table></td>
        </tr>
        <%
do while not rs.eof 
%>
        <tr> 
          <td colspan="4"><%=rs("eid")%>.<a href="events_show.asp?eid=<%=rs("eid")%>"><%=rs("name")%></a>&nbsp;</td>
        </tr>
        <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
rs.movefirst
%>
        <tr> 
          <td height="35" colspan="6"> <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td> <table border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr> 
                      <%
zj=1
for zi=1 to rs.pagecount
%>
                      <td width="50"><a href="events.asp?page=<%=zj%>">|&nbsp;<%=zj%>&nbsp;|</a></td>
                      <%
zj=zj+1
next
%>
                    </tr>
                  </table></td>
                <td width="20%" align="left"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td> <div align="center"> Record :<font color=red><b><%=rs.recordcount%></b></font> Page <font color=red><b><%=pagecount%></b></font> of <font color=red><b><%=rs.pagecount%></b></font> </div></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <%
rs.close
set rs=nothing
%> </td>
    <td width="181">&nbsp;</td>
  </tr>
</table>
<!--#include file="Copyright.asp"-->