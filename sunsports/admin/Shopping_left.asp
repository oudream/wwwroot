<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="185" border="0" align="center" cellpadding="0" cellspacing="0">
<%
if session("user_uid")<>"" then
%>
<%
if session("user_adminlevel")=1 then
%>
<!-- 
   ----------------------------------------------   ----------------------------------------------------
   														logined level1																
   ----------------------------------------------   ----------------------------------------------------
-->
  <tr> 
    <td>&nbsp;</td>
	    
    <td align="center"><TABLE height=100% width=100% border=0>
        <TBODY>
          <TR> 
            <TD vAlign=top align=left> <TABLE width=100% height="30" border=0>
                <TBODY>
                  <TR> 
                    <TD width="50%" height="40"><strong><a href="left.asp" target="leftFrame"><font color="#FF0000">SPORTS</font></a></strong></TD>
                    <TD width="50%"><strong><a href="Shopping_left.asp" target="leftFrame"><font color="#0066FF">PRODUCTS</font></a></strong></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#f5f5f5">
                <TBODY>
                  <TR>
                    <TD height="25"><strong>Sort &amp; Products</strong></TD>
                  </TR>
                  <TR> 
                    <TD width=118 height="26"><a href="sort_c.asp" target="mainFrame">Sort 
                      Add </a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"> <a href="sort_v.asp" target="mainFrame">Sort 
                      Edit/Del</a></TD>
                  </TR>
                  <TR> 
                    <TD width=118 height="25"><a href="subs_c.asp" target="mainFrame">Product 
                      Add </a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"> <a href="subs_v.asp" target="mainFrame">Product 
                      Edit/Del</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5">
                <TBODY>
                  <TR>
                    <TD height="25"><strong>Payment</strong></TD>
                  </TR>
                  <TR> 
                    <TD width=118 height="25"><a href="payment_c.asp" target="mainFrame">Payment 
                      Add </a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"> <a href="payment_v.asp" target="mainFrame">Payment 
                      Edit/Del</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5">
                <TBODY>
                  <TR>
                    <TD height="25"><strong>Payout</strong></TD>
                  </TR>
                  <TR> 
                    <TD width=118 height="25"><a href="payout_c.asp" target="mainFrame">Payout 
                      Add</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"> <a href="payout_v.asp" target="mainFrame">Payout 
                      Edit/Del</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5">
                <TBODY>
                  <TR>
                    <TD height="25"><strong>Shopper</strong></TD>
                  </TR>
                  <TR> 
                    <TD width=118 height="25"><a href="shopping_user_v.asp" target="mainFrame">User 
                      Logined View</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5">
                <TBODY>
                  <TR>
                    <TD height="25"><strong>Orders</strong></TD>
                  </TR>
                  <TR> 
                    <TD width=118 height="25"><a href="shopping_order_v.asp" target="mainFrame">Order 
                      View</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              
            </TD>
          </TR>
        </TBODY>
      </TABLE> </td>
        <td></td>
  </tr>
<%
end if
%>
<!-- 
   ----------------------------------------------   ----------------------------------------------------
   														loginout																
   ----------------------------------------------   ----------------------------------------------------
-->
  <tr>
    <td>&nbsp;</td>
    <td height="50" align="center" valign="middle">
          <FORM action=login.asp method="post" id="form_logout" name="form_logout" target="_parent">
            <INPUT type=submit value="logout" id="option" name="option">
          </FORM>
	</td>
    <td>&nbsp;</td>
  </tr>
<%
else response.Redirect("error.asp?../err=001")
end if
%>
</table>
</body>
</html>