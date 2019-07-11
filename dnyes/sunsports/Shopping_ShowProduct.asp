<!--#include file="Top_sun.asp"-->
<table width="776" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="181" valign="top"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td>&nbsp;</td>
        </tr>
      </table>
      <!--#include file="Shopping_Sort.asp" -->	</td>
    <td width="20">&nbsp; </td>
    <td valign="top"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="27">&nbsp;</td>
        </tr>
      </table>
      <%
subs_id=request("subs_id")
sql="select * from subs where subs_id=" & subs_id
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1 
%>
      <table width="360" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td align="center"><img src="<%=rs("photo")%>" width="360" height="360"></td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="50" align="center"><b><%=rs("name")%></b></td>
        </tr>
      </table>
      <table width="360" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td>Product ID:</td>
          <td><%=rs("code")%>&nbsp;</td>
        </tr>
        <tr> 
          <td>Price:</td>
          <td>$ <%=rs("price")%></td>
        </tr>
        <tr> 
          <td>Brand:</td>
          <td><%=rs("brand")%>&nbsp;</td>
        </tr>
        <tr> 
          <td>Size:</td>
          <td><%=rs("size")%>&nbsp;</td>
        </tr>
        <tr> 
          <td>Packing:</td>
          <td><%=rs("packaging")%>&nbsp;</td>
        </tr>
        <tr> 
          <td>Made in:</td>
          <td><%=rs("madein")%>&nbsp;</td>
        </tr>
        <tr> 
          <td>Add date:</td>
          <td><%=rs("putdate")%>&nbsp;</td>
        </tr>
        <tr> 
          <td colspan="2"><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td>Procuct Description:<%=rs("content")%></td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
              </tr>
            </table></td>
        </tr>
        <tr align="center"> 
          <td colspan="2"> <a href="Shopping_Buy.asp?subs_id=<%=rs("subs_id")%>" target="_blank">[ ORDER NOW ]</a> </td>
        </tr>
      </table>
      <%
rs.close
set rs=nothing
%>
      <br> </td>
    <td width="20" rowspan="2" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="24">&nbsp;</td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td>&nbsp;</td>
          <td height="550" background="images/10x1_blue.gif">&nbsp;</td>
        </tr>
      </table></td>
    <td width="181" rowspan="2" valign="top"> 
	<!--#include file="User_l.asp" -->
	 <br> <!--#include file="Shopping_Search.asp" -->
      <table width="100%" border="0" cellpadding="5" cellspacing="0">
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td><font color="#1F4A71"><strong>Questions?</strong></font><br>
            Call us please , SUNSPORTS (65)63448344 , 24 hours a day, 7 days a 
            week.</td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table>
      <br>
    </td>
  </tr>
  <tr> 
    <td colspan="3" valign="top">&nbsp;</td>
  </tr>
</table>
<!--#include file="Copyright.asp"-->