<!--#include file="adminconn.asp" -->
<%
orders_id=request("orders_id")
if orders_id<>"" then 
osql="select * from orders where orders_id="&orders_id
else
osql="select * from orders order by orders_id desc"
end if
Set ors=Server.CreateObject("ADODB.RecordSet") 
ors.open osql,conn,1,1

subs_id=ors("subs_id")
sql="select * from subs where subs_id="& subs_id
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr align="center"> 
    <td height="50" colspan="3">order no. : <%=orders_id%></td>
  </tr>
  <tr> 
    <td width="50">&nbsp;</td>
    <td><table width="90%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#6184A3">
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="left"><img src="../<%=rs("pic")%>" width="90" height="120"></td>
          <td width="75%"><p>Name: <%=rs("name")%><br>
              Brand: <%=rs("brand")%> <br>
              <%if len(rs("content")) > 100 then
				response.Write(left(rs("content"),98)&"..")
				else
				response.Write(rs("content"))
				end if%>
              <br>
            </p>
            <p>Size: <%=rs("size")%><br>
              Price: $<%=rs("price")%><br>
            </p></td>
        </tr>
        <%
rs.close
set rs=nothing
%>
        <tr bgcolor="#FFFFFF"> 
          <td>Size:</td>
          <td><%=ors("size")%>&nbsp; </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td>Quantity:</td>
          <td><%=ors("count")%>&nbsp; </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td>How to pay:</td>
          <td>
            <%
			if not isnull(ors("payment_id")) then
		  	psql="select * from payment where payment_id=" & ors("payment_id") 
			Set prs=Server.CreateObject("ADODB.RecordSet") 
			prs.Open psql,conn,1,1 
			if not prs.eof then
			response.Write(prs("name"))
			end if
			prs.close
			set prs=nothing
			end if
		  	%>&nbsp; 
		  </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td>How to get:</td>
          <td>
            <%
			if not isnull(ors("payout_id")) then
		  	psql="select * from payout where payout_id=" & ors("payout_id") 
			Set prs=Server.CreateObject("ADODB.RecordSet") 
			prs.Open psql,conn,1,1 
			if not prs.eof then
			response.Write(prs("name"))
			end if
			prs.close
			set prs=nothing
			end if
		  	%>&nbsp; 
		  </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td>Content:</td>
          <td><%=ors("Content")%>&nbsp; </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="2">&nbsp;</td>
        </tr>
        <tr align="center" bgcolor="#FFFFFF"> 
          <td colspan="2"><strong>Shopping User Info</strong></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td>Full Name:</td>
          <td><%=ors("name")%>&nbsp; </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td>TEL:</td>
          <td><%=ors("tel")%>&nbsp; </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td>ZIP:</td>
          <td><%=ors("zip")%>&nbsp; </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td>Address:</td>
          <td><%=ors("contact")%>&nbsp; </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td>Email:</td>
          <td><%=ors("email")%>&nbsp; </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="2">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="2"><table width="100%" height="35" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td align="center"><strong><a href="Shopping_Order_V.asp?page=<%=request("page")%>">go 
                  back</a></strong></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
    <td width="50">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<%
ors.close
set ors=nothing
%>
