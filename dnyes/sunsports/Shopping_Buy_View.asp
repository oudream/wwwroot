<script language="JavaScript">

function checkform_buy()
{
	if (form_buy.count.value.length == 0) {
		alert("Please enter a valid Quantity. ");
		form_buy.count.focus();
		return false;
		}
	if(! isNumber(form_buy.count.value)) {
		alert('The "Quantity" field makes up by "0~9"');
		form_buy.count.focus();
		return false;
		}
	if (form_buy.payment.value.length == 0) {
		alert("Please enter a valid 'How to pay'. ");
		form_buy.payment.focus();
		return false;
		}
	if (form_buy.payout.value.length == 0) {
		alert("Please enter a valid 'How to get'. ");
		form_buy.payout.focus();
		return false;
		}
	if (form_buy.uname.value.length == 0) {
		alert("Please enter a valid 'Full Name'. ");
		form_buy.uname.focus();
		return false;
		}
	if (form_buy.tel.value.length == 0) {
		alert("Please enter a valid 'Telphone'. ");
		form_buy.tel.focus();
		return false;
		}
	if (form_buy.zip.value.length == 0) {
		alert("Please enter a valid 'Zip'. ");
		form_buy.zip.focus();
		return false;
		}
	if (form_buy.contact.value.length == 0) {
		alert("Please enter a valid 'Address'. ");
		form_buy.contact.focus();
		return false;
		}
	if(form_buy.email.value.length > 0){
	if(! isMail(form_buy.email.value)) {
		alert('Please enter a valid EMAIL address');
		form_buy.email.focus();
		return false;
		}}
	return true;
}

function isLetter(name)
{
var letter=" _0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
	if(name.length == 0)
		return false;
	for(i = 0; i < name.length; i++) {
		if(letter.indexOf(name.charAt(i)) == -1)
			return false;
	}
	return true;
}

function isNumber(name)
{
	for(i = 0; i < name.length; i++) {
		if(name.charAt(i) < "0" || name.charAt(i) > "9")
			return false;
	}
	return true;
}

function isEnglish(name)
{
	for(i = 0; i < name.length; i++) {
		if(name.charCodeAt(i) > 128)
			return false;
	}
	return true;
}

function isMail(name)
{
	if(! isEnglish(name))
		return false;
	i = name.indexOf("@");
	j = name.lastIndexOf("@");
	if(i == -1)
		return false;
	if(i != j)
		return false;
	if(i == name.length)
		return false;
	return true;
}

</script>
<!--#include file="Top_sun.asp"-->
<%
if request("option")="del" then
conn.execute "delete * from orders where orders_id=" & request("orders_id")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Del the order is complete. ');</SCRIPT>")
end if
%>
<table width="776" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="8">&nbsp; </td>
    <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="24">&nbsp;</td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#6184A3">
        <tr bgcolor="#FFFFFF"> 
          <td height="30" colspan="6" align="center"><strong>Your Order List</strong></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="70" height="30" align="center"><strong>ORDER NO.</strong></td>
          <td align="center"><strong>NAME</strong></td>
          <td width="70" align="center"><strong>UNIT PRICE</strong></td>
          <td width="70" align="center"><strong>QUANTITY</strong></td>
          <td width="70" align="center"><strong>TOTAL</strong></td>
          <td width="70" align="center"><strong>OPERATION</strong></td>
        </tr>
        <%
ouifdata="no"		
osql="select * from orders where pid="&session("user_pid")
Set ors=Server.CreateObject("ADODB.RecordSet")
ors.open osql,conn,1,1

do while not ors.eof
ouifdata="yes"
subs_id=ors("subs_id")
sql="select * from subs where subs_id="& subs_id
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.open sql,conn,1,1
%>
        <tr bgcolor="#FFFFFF"> 
          <td height="30" align="center"><%=ors("orders_id")%>&nbsp;</td>
          <td align="center"><a href="Shopping_ShowProduct.asp?subs_id=<%=rs("subs_id")%>" target="_blank"><%=rs("name")%></a>&nbsp;</td>
          <td align="center">$ <%=rs("price")%>&nbsp;</td>
          <td align="center"><%=ors("count")%></td>
          <td align="center">$ <%=ccur(rs("price"))*ccur(ors("count"))%>&nbsp;</td>
          <td align="center"><a href="Shopping_Buy_Edit.asp?orders_id=<%=ors("orders_id")%>" target="_blank">Edit</a> 
            &nbsp; <a href="Shopping_Buy_View.asp?option=del&orders_id=<%=ors("orders_id")%>">Del</a></td>
        </tr>
        <%
rs.close
set rs=nothing
ors.movenext
loop
ors.close
set ors=nothing
%>
<%
if ouifdata="no" then
%>
        <tr bgcolor="#FFFFFF"> 
          <td height="30" colspan="6" align="center"><strong> No order Now!</strong></td>
        </tr>
<%
end if
%>
        <tr bgcolor="#FFFFFF"> 
          <td height="30" colspan="6" align="center">&nbsp;</td>
        </tr>
      </table></td>
    <td width="20" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="24">&nbsp;</td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td>&nbsp;</td>
          <td height="450" background="images/10x1_blue.gif">&nbsp;</td>
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