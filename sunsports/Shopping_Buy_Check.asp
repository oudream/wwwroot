<!--#include file="Top_sun.asp"-->
<table width="776" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="181" valign="top">
<!--#include file="Shopping_Sort.asp" -->	</td>
    <td width="20">&nbsp; </td>
    <td valign="top">
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
	<table width="360" border="0" align="center" cellpadding="5" cellspacing="1">
        <tr bgcolor="#FFFFFF"> 
          <td width="120"><img src="<%=rs("pic")%>" width="120" height="120"></td>
          <td><p><b><%=rs("name")%></b><br>
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
          <td colspan="2"><font color="#FF0000">#</font> Check out your order 
            form please . If your want some change, click Edit below . If it is 
            OK , click Close and we will deal with the order as soon as possible 
            .Thank you!</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="2">The Size You Want:<%=ors("size")%> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="2">Quantity:<%=ors("count")%> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="2">How To Pay:
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
		  	%>
		  </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="2">How To Get:
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
		  	%>
		  </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="2">Content:<%=ors("Content")%> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="2">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="2">Full Name:<%=ors("name")%>&nbsp; </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="2">Telphone Number:<%=ors("tel")%>&nbsp; </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="2">ZIP/Postal Code:<%=ors("zip")%>&nbsp; </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="2">Address:<%=ors("contact")%>&nbsp; </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="2">Email:<%=ors("email")%>&nbsp; </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="2">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="2"><table width="100%" height="35" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="51%" align="center"><A href="javascript:parent.close();"><strong>Close</strong></A></td>
                <td width="49%" align="center"><strong><a href="Shopping_Buy_Edit.asp?orders_id=<%=ors("orders_id")%>">Edit</a></strong></td>
              </tr>
            </table></td>
        </tr>
      </table>
<%
ors.close
set ors=nothing
%>	  
	   </td>
    <td width="20" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="24">&nbsp;</td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td>&nbsp;</td>
          <td height="500" background="images/10x1_blue.gif">&nbsp;</td>
        </tr>
      </table></td>
    <td width="181" valign="top">
<!--#include file="User_l.asp" -->
	 <br> <!--#include file="Shopping_Search.asp" -->
	 <br> 	
	
	</td>
  </tr>
</table>
<!--#include file="Copyright.asp"-->