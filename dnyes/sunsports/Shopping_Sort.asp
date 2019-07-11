<table width="170" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr>
                
          <td><a href="Shopping_sort_view.asp"><b><font color="#1F4A71">SUNSPORTS PRODUCTS 
            </font></b></a></td>
              </tr>
            </table>
<%
sql="select * from sort" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 
do while not rs.eof
%>
	<table width="100%" border="0" cellspacing="3" cellpadding="0">
        <tr> 
          <td width="5" valign="top"><img src="images/sun_l_sp.gif" width="5" height="10" border="0"></td>
          <td><a href="Shopping_sort_view.asp?sort_id=<%=rs("sort_id")%>" class="b"><%=rs("name")%></a></td>
        </tr>
      </table>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
