<html>
<head>
</head>
<body bgcolor=#c1f7d8>
<p align=center><font size=5> �������� </font></p>
<form action=sendmessage.asp method=post>
<table align=center border=1 cellpadding=1 cellspacing=1 width=300>
  <tr>
   <td height=23>&nbsp;����:</td>
   <td height=23><input type=hidden size=20 name=toname value=<%=request.querystring("toname")%>>
   <%=request.querystring("toname")%></td>
  </tr>
  <tr>
   <td height=23>&nbsp;����:</td>
   <td height=23><input type=hidden name=name value=<%=session("name")%>><%=session("name")%></td>
  </tr>
  <tr>
   <td height=19>&nbsp;����:</td>
   <td height=19>&nbsp;</td>
  </tr>
  <tr>
   <td colspan=2> &nbsp;<textarea id=textarea1 name=content rows=6 cols=35></textarea>
   <p align=center>
   <input id=submit1 name=submit1 type=submit value=��������>
   <input id=reset1 name=reset1 type=reset value=������д>
   </td>
  </tr>
</table>
</form>
</body>
</html>







