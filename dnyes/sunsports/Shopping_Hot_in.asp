<table width="170" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="35"><b><font color="#1F4A71">What's Hot !</font></b></td>
  </tr>
</table>
<table width="160" border="0" align="center" cellpadding="0" cellspacing="0">
  <%
i=0
osql=" SELECT subs_id ,sum(count) as sumcount FROM orders GROUP BY subs_id ORDER BY Sum(count) DESC"
Set ors=Server.CreateObject("ADODB.RecordSet") 
ors.open osql,conn,1,1 
do while not ors.eof


sql="select * from subs where subs_id="&ors("subs_id")
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 
%>
  <tr> 
    <td> <b><%=rs("name")%></b><br>
      Price:<font color="#FF0000"> $ <%=rs("price")%></font> Ordered <%=ors("sumcount")%> Times<br>
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
      <br>
      <br>
				</td>
  </tr>
  <%
rs.close
set rs=nothing
ors.movenext
i=i+1                                                                     
if i>=5 then exit do                                                           
loop
ors.close
set ors=nothing
%>
</table>
