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
          <td>&nbsp;</td>
        </tr>
      </table>
      <%
subs_id=request("subs_id")
sql="select * from subs where subs_id="& subs_id
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
%>
	<table width="100%" border="0" cellpadding="0" cellspacing="1">
        <tr> 
          <td><img src="<%=rs("pic")%>" width="120" height="120"></td>
          <td><p><b><%=rs("name")%></b><br>
              Brand: <%=rs("brand")%> <br>
              Size: <%=rs("size")%><br>
              Price: <font color="#FF0000">$<%=rs("price")%></font><br>
              Description: 
              <%if len(rs("content")) > 100 then
				response.Write(left(rs("content"),98)&"..")
				else
				response.Write(rs("content"))
				end if%>
              <br>
              <br>
            </p></td>
        </tr>
        <%
rs.close
set rs=nothing
%>
        <tr> 
          <td colspan="2" align="center"><strong>Order Form</strong></td>
        </tr>
        <tr>
          <td height="25" colspan="2"><font color="#FF0000">#</font> We will delivering 
            the products according this form.<br>
            <br>
            <font color="#FF0000">*</font><b>Required fields</b>. </td>
        </tr>
        <FORM action="Shopping_Buy_Save.asp" method="post" name="form_buy" id="form_buy" onSubmit="return checkform_buy();">
          <tr> 
            <td>The Size You Want:</td>
            <td><input name="size" type="text" id="size" size="20" maxlength="50" class="form">
              ( e.g., Size XXL )</td>
          </tr>
          <tr> 
            <td>Quantity:</td>
            <td><input name="count" type="text" id="count" size="10" maxlength="10" class="form"> 
              <font color="#FF0000">* </font></td>
          </tr>
          <tr> 
            <td>How To Pay:</td>
            <td> <select name="payment" id="payment" class="display_dropdown">
                <option value="" selected>Choose one please</option>
                <option>----------</option>
                <%
sqlp="select * from payment"
set rsp=server.createobject("ADODB.Recordset")
rsp.open sqlp,conn,1,1
while not rsp.eof
%>
                <option value="<%=rsp("payment_id")%>"><%=rsp("name")%></option>
                <%
rsp.movenext
wend
rsp.close
set rsp=nothing
%>
              </select> <font color="#FF0000">* </font> </td>
          </tr>
          <tr> 
            <td>How To Get:</td>
            <td> <select name="payout" id="payout" class="display_dropdown">
                <option value="" selected>Choose one please</option>
                <option>----------</option>
                <%
sqlp="select * from payout"
set rsp=server.createobject("ADODB.Recordset")
rsp.open sqlp,conn,1,1
while not rsp.eof
%>
                <option value="<%=rsp("payout_id")%>"><%=rsp("name")%></option>
                <%
rsp.movenext
wend
rsp.close
set rsp=nothing
%>
              </select> <font color="#FF0000">* </font> </td>
          </tr>
          <tr> 
            <td>Content:</td>
            <td><textarea name="Content" cols="50" rows="8" class="form" id="Content"></textarea></td>
          </tr>
          <tr> 
            <td colspan="2">&nbsp;</td>
          </tr>
          <tr> 
            <td>Full Name:</td>
            <td><input name="uname" type="text" id="uname" size="30" maxlength="20" value="<%=session("user_name")%>" class="form"> 
              <font color="#FF0000">* </font></td>
          </tr>
          <tr> 
            <td>Telphone Number:</td>
            <td><input name="tel" type="text" id="tel" size="30" maxlength="20" value="<%=session("user_tel")%>" class="form"> 
              <font color="#FF0000">* </font></td>
          </tr>
          <tr> 
            <td>ZIP/Postal Code:</td>
            <td><input name="zip" type="text" id="zip" size="10" maxlength="10" value="<%=session("user_zip")%>" class="form"> 
              <font color="#FF0000">* </font></td>
          </tr>
          <tr> 
            <td>Address:</td>
            <td><input name="contact" type="text" id="contact" size="50" maxlength="100" value="<%=session("user_contact")%>" class="form"> 
              <font color="#FF0000">* </font></td>
          </tr>
          <tr> 
            <td>Email:</td>
            <td><input name="email" type="text" id="email" size="30" maxlength="30" value="<%=session("user_email")%>" class="form"></td>
          </tr>
          <tr> 
            <td height="80" colspan="2"><font color="#FF0000">#</font> To ensure 
              that your orders are processed quickly, all information you provide 
              must be accurate and verifiable.</td>
          </tr>
          <tr> 
            <td colspan="2"> <input type="hidden" name="subs_id" id="subs_id" value=<%=subs_id%>> 
              <input type="hidden" name="option" id="option" value="add"> <input type="submit" name="Submit" value="Order Now !" class="button"> 
            </td>
          </tr>
        </form>
        <tr> 
          <td>&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
      </table> </td>
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