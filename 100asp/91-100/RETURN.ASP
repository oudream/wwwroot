<html>
<head>
<title></title>
</head>
<body bgcolor=#c1f7d8>
<p align=center><font size=5>ªÿ∏¥¡Ù—‘</font></p>
<form action=sendmessage.asp method=post>
<table align=center border=1 cellpadding=1 cellspacing=1 width=300>
   <tr>
     <td height=23>&nbsp; ªÿ∏¥:</td>
     <td height=23><input type=hidden size=20 name=toname value=<%=request.querystring("toname")%>>
                   <%=request.querystring("toname")%></td>
   </tr>
   <tr>
      <td height=23>&nbsp;–’√˚:</td>
      <td height=23><input type=hidden name=name value=<%=session("name")%>><%=session("name")%></td>
   </tr>
   <tr>
      <td height=19>&nbsp;¡Ù—‘:</td>
      <td height=19>&nbsp;</td>
   </tr>
   <tr>
   <td colspan=2> &nbsp;<textarea name=content rows=6 cols=35></textarea>
   <p align=center>
   <input type=submit value=∑¢ÀÕ¡Ù—‘>
   <input type=reset value=÷ÿ–¬ÃÓ–¥></td>
   </tr>
</table>
</form>
</body>
</html>