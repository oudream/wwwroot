<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<%on error resume next%>
<script language="JavaScript">
function CheckForm()
{
    if (document.payment.paymenttype.value.length == 0) {
		alert("请选择您的支付方式.");
		document.payment.paymenttype.focus();
		return false;
	}
	return true;
}
</script>
<%
'判断是否正常购买，如果不是，需要登陆
if session("y_it_uid")="" then 
response.redirect "error.asp?err=011"
response.end
end if

'判断购物车是否为空
ProductList = Session("ProductList")
if productlist<>"" then
sql="select * from subs where bookbm in ("&productlist&") order by bookbm"
Set rs = conn.Execute( sql )
else
response.redirect "error.asp?err=007"
response.end
end if
%> 
 <br>
<table border="0" cellspacing="1" cellpadding="0" align="center" width="500" bgcolor="#285278">
  <tr bgcolor="#BFCEDD"> 
    <td colspan="4" height="22">&nbsp;&nbsp;<%=session("y_it_uid")%>您选购的商品清单如下： 
    </td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td width="35%" height="22"> 
      <div align="center">商 品 名 称</div>
    </td><td width="25%" height="22"> 
      <div align="center">供 应 商</div>
    </td>
    <td   width="15%"> 
      <div align="center">年 限</div>
    </td>
    <td   width="25%"> 
      <div align="center">价 格</div>
    </td>
  </tr>
<%
   Sum = 0
   While Not rs.EOF
     Quatity = CInt( Request( "Q_" & rs("bookbm")) )
     If Quatity <= 0 Then 
        Quatity = CInt( Session(rs("bookbm")) )
        If Quatity <= 0 Then Quatity = 1
     End If
     Session(rs("bookbm")) = Quatity

     if rs("ifdisc")=true then
     Sum = Sum + csng(rs("price")) * Quatity * session("y_it_userdiscount")
     else
     Sum = Sum + csng(rs("price")) * Quatity
     end if

     Sum=FormatNumber(Sum,2) 
session("sum")=sum
%> 
  <tr bgcolor="#FFFFFF"><%
if instr(1,rs("subsname"),".",1)<>0 then
temp=Session("domain")&rs("subsname")
else
temp=rs("subsname")
end if
%> 
    <td height="22">&nbsp;&nbsp;<%=temp%></td> 
    <td height="22">&nbsp;&nbsp;<%=rs("gys")%></td>
    <td width="15%" align="center"><%=Quatity%> </td>
    <td width="15%" align="right">&nbsp;&nbsp;
<%if rs("ifdisc")=true then%>
<%=FormatNumber(csng(rs("price"))*Quatity*session("y_it_userdiscount"),2)%>
<%else%>
<%=FormatNumber(csng(rs("price"))*Quatity,2)%>
<%end if%></td>
  </tr>
<%
rs.MoveNext
Wend
rs.close
set rs=nothing
%> 
  <tr bgcolor="#FFFFFF"> 
    <td colspan="4" height="25"> 
      <div align="right">总价格  =  <%=Sum%></div>
    </td>
  </tr>
</table>
<br>
<br>
<table border="0" cellspacing="1" cellpadding="0" align="center" width="500" bgcolor="#285278">
  <form name="payment" action="dopayment.asp" method="POST"  onSubmit="return CheckForm();">
    <tr bgcolor="#FFFFFF"> 
      <td height="30" width="20%">&nbsp;您的帐号：</td>
      <td width="30%">&nbsp;<%=session("y_it_uid")%></td>
      <td width="20%">&nbsp;您的预付款：</td>
      <td width="30%" align="right">&nbsp;<%
sqlt="select prepay from user where uid='"&session("y_it_uid")&"'"
set rst=server.createobject("adodb.recordset") 
rst.open sqlt,conn,1,1
response.write rst("prepay") & " RMB"
rst.close
set rst=nothing
%></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="30">&nbsp;支付方式</td>
      <td colspan="3">&nbsp; 
        &nbsp;&nbsp; </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="35" colspan="4"> 
        <div align="center"><INPUT onclick=history.go(-1) type=button value="<<上一步" name=button>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
          <input type="submit" name="Submit22" value="下一步>>">
		  <input type=hidden value=<%=paymenttype%> name=paymenttype>
        </div>
      </td>
    </tr>
  </form>
</table>
<br>
<%
if trim(request("paymenttype"))<>""then
response.Redirect("dopayment.asp?paymenttype="&request("paymenttype"))
response.End()
else
response.Redirect("index.asp")
response.End()
end if
%>
