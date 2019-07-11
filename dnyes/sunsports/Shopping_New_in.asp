<table width="170" border="0" align="center" cellpadding="0" cellspacing="0">
<tr><td colspan="2"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="3">
        <tr> 
          <td><a href="Shopping_sort_view.asp"><b><font color="#1F4A71">What's 
            New !</font></b></a></td>
        </tr>
      </table></td></tr>
  <%
i=0
sql="select * from subs order by subs_id desc"
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 

if not isempty(request("page")) then   
pagecount=cint(request("page"))   
else   
pagecount=1
end if

rs.pagesize=4
rs.AbsolutePage=pagecount

do while not rs.eof
%>
  <tr> 
    <td>
	  <B><%=rs("name")%></B><br>
      Price:<font color="#FF0000"> $ <%=rs("price")%></font> 
	  <br> <a href="Shopping_Buy.asp?subs_id=<%=rs("subs_id")%>">&lt;BUY&gt;</a> 
      <br> 
	</td>
    <td width="60"><a href="Shopping_ShowProduct.asp?subs_id=<%=rs("subs_id")%>" target="_blank"><img src="<%=rs("pic")%>" width="60" height="60" border="0"></a></td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <%
				if len(rs("content")) > 100 then
				response.Write(left(rs("content"),98)&"..")
				else
				response.Write(rs("content"))
				end if
	 %>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="5"><img src="1.gif" width="1" height="1"></td>
        </tr>
      </table>
      <table width="70%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="1" bgcolor="#6184A3"><img src="1.gif" width="1" height="1"></td>
        </tr>
      </table> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="25">&nbsp;</td>
        </tr>
      </table> </td>
  </tr>
  <%
rs.movenext
i=i+1                                                                     
if i>=rs.pagesize then exit do                                                           
loop
%>
</table>
