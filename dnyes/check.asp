<script language="JavaScript">
function CheckForm()
{
	if (document.check.paymenttype.value.length == 0) {
		alert("请选择支付方式");
		document.check.paymenttype.focus();
		return false;
	}
		return true;
}
</script>
<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<%
Sub PutToShopBag( bookbm, ProductList )
   If Len(ProductList) = 0 Then
      ProductList = "'" & bookbm & "'"
ElseIf InStr( ProductList, bookbm ) <= 0 Then
      ProductList = ProductList & ", '" & bookbm & "'"
   End If
End Sub
%>
<%
'如果购买车为空，转入错误界面
ProductList = Session("ProductList")
If Len(ProductList) = 0 Then
   Response.Redirect "error.asp?err=007"
   response.end
end if

'取消要购买物品的处理

Quatityt=""
subslist=""

If Request("cmdShow") = "Yes" Then
   ProductList = ""
   Products = Split(Request("bookbm"), ", ")
   For I=0 To UBound(Products)
      PutToShopBag Products(I), ProductList
   Next
   Session("ProductList") = ProductList
End If

%>


<%
sqlp2 = "SELECT * FROM user where uid= '" &session("y_it_uid")& "'"
set rsp2 = Server.CreateObject("ADODB.Recordset")
rsp2.open sqlp2,conn,1,1
zcontactz=rsp2("contactz")
zemail=rsp2("email")
ztel=rsp2("tel")
zdizhiz=rsp2("dizhiz")
rsp2.close
set rsp2=nothing
%>  


<%


if productlist<>"" then
Set rsc=Server.CreateObject("ADODB.RecordSet") 
sqlc="select * from subs where bookbm in ("&productlist&") order by bookbm"
rsc.open sqlc,conn,1,1
else
   Response.Redirect "error.asp?err=007"
end if
%>
<html>
<head>
<title><%=Application("y_it_itname")%>-商品确认</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0">
<!--#include file="top.asp"-->
<table width="776" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="10" background="images/dnyes-bg-left-and-right.gif"><img src="images/dnyes-bg-left-and-right.gif" width="10" height="1"></td>
    <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td colspan="5"><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="1" background="images/2x2.gif"><img src="images/1x2-black.gif" width="1" height="2"></td>
                <td><img src="images/1-x.gif" width="754" height="56"></td>
                <td width="1" background="images/1x2-black.gif"><img src="images/1x2-black.gif" width="1" height="2"></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td width="229" valign="top" background="images/left-228x5.gif"> <table width="229" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="229" height="53" background="images/left-1b.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="10%" height="18" valign="bottom">&nbsp; </td>
                      <td width="8%" valign="bottom"><img src="images/gogo.gif" width="6" height="15"></td>
                      <td width="82%" valign="bottom"><font color="#FFFFFF"><b>客 
                        户 登 录</b></font></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td valign="top"> <table width="228" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td background="images/left-1-left-2b.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td width="21" valign="top"><img src="images/left-1-left-1b.gif" width="21" height="190"></td>
                            <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                                <tr> 
                                  <td> 
                                    <!--#include file="userlogin.asp"-->
                                  </td>
                                </tr>
                                <tr> 
                                  <td> 
                                    <!--#include file="userhelp.asp" -->
                                  </td>
                                </tr>
                              </table></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td><img src="images/left-1-left-3b.gif" width="229" height="12"></td>
                    </tr>
                  </table></td>
              </tr>
            </table> 
            <table width="229" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="229" height="28" background="images/left-2.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="10%" height="18" valign="bottom">&nbsp; </td>
                      <td width="8%" valign="bottom"><img src="images/gogo.gif" width="6" height="15"></td>
                      <td width="82%" valign="bottom"><font color="#FFFFFF"><b>客 
                        户 中 心</b></font></td>
                    </tr>
                  </table></td>
              </tr>
            </table>
            <TABLE width=70% border=0 cellPadding=0 cellSpacing=0>
              <TBODY>
                <TR> 
                  <TD height="3" 
                vAlign=top 
                style="BORDER-left: #000000 1px solid"><img src="1.gif" width="1" height="1"></TD>
                </TR>
              </TBODY>
            </TABLE>
            <table width="228" height="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="1" background="images/1x2-black.gif"><img src="images/1x2-black.gif" width="1" height="2"></td>
                <td valign="top"> 
                  <!--#include file="support.asp" -->
                </td>
              </tr>
            </table></td>
          <td width="7">&nbsp;</td>
          <td width="513" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
                <td colspan="5" height="88"> <TABLE cellSpacing=0 cellPadding=0 width=100% border=0>
                    <TBODY>
                      <TR> 
                        <TD width="34%" vAlign=center style="BORDER-left: #CCCCCC 1px solid; BORDER-RIGHT: #CCCCCC 1px solid; BORDER-BOTTOM: #CCCCCC 1px solid;BORDER-TOP: #CCCCCC 1px solid"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td width="12" height="25" background="images/host-2.gif">&nbsp;</td>
                              <td valign="bottom" background="images/host-2.gif">购 
                                买 产 品</td>
                            </tr>
                          </table>
                          <table width="75%" border="0" cellpadding="0" cellspacing="0">
                            <tr>
                              <td height="6"><img src="1.gif" width="1" height="1"></td>
                            </tr>
                          </table>
                          <table width="498" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#cccccc">
                            <form action="Check.asp" method="POST" name="check" onSubmit="return CheckForm();">
                              <tr bgcolor="#f5f5f5"> 
                                <td height="35" colspan="2"><%=zcontactz%> 您选购的商品清单如下：</td>
                              </tr>
                              <%
Sum = 0
   While Not rsc.EOF
     Quatity = CInt( Request( "Q_" & rsc("bookbm")) )
     If Quatity <= 0 Then 
      	Quatity = CInt( Session(rsc("bookbm")) )
        If Quatity <= 0 Then Quatity = 1
     End If


	if len(Quatityt)=0 then
		Quatityt=Quatity
		else
		Quatityt=Quatityt&","&Quatity
	end if

	if len(gyslist)=0 then
		gyslist=rsc("gys")
	else
		gyslist=gyslist&","&rsc("gys")
	end if

	if len(subslist)=0 then
		if instr(1,rsc("subsname"),".",1)<>0 then
		subslist=Session("domain")&rsc("subsname")
		else
		subslist=rsc("subsname")
		end if
	else
		if instr(1,rsc("subsname"),".",1)<>0 then
		subslist=subslist&","&Session("domain")&rsc("subsname")
		else 
		subslist=subslist&","&rsc("subsname")
		end if
	end if


     Session(rsc("bookbm")) = Quatity
     if rsc("ifdisc")=true then
     Sum = Sum + csng(rsc("price")) * Quatity * session("y_it_userdiscount")
     else
     Sum = Sum + csng(rsc("price")) * Quatity
     end if
     Sum=FormatNumber(Sum,2) 
%>
                              <input type="hidden" name="bookbm" value="<%=rsc("bookbm")%>">
                              <input type="hidden" name="subs" value="<%=rsc("subsname")%>">
                              <%
	if instr(1,rsc("subsname"),".",1)<>0 then
	temp=Session("domain")&rsc("subsname")
	else
	temp=rsc("subsname")
	end if
%>
                              <tr> 
                                <td width="28%" height="25" bgcolor="#f5f5f5">商 
                                  品 ID</td>
                                <td width="72%" bgcolor="#FFFFFF"><%=rsc("id")%><%session("subid")=rsc("id")%></td>
                              </tr>
                              <tr> 
                                <td height="25" bgcolor="#f5f5f5"><font color="#000000">商 
                                  品 名 称</font></td>
                                <td bgcolor="#FFFFFF"><%=temp%></td>
                              </tr>
                              <tr> 
                                <td height="25" bgcolor="#f5f5f5"><font color="#000000">供 
                                  应 商</font></td>
                                <td bgcolor="#FFFFFF"><%=rsc("gys")%></td>
                              </tr>
                              <tr> 
                                <td height="25" bgcolor="#f5f5f5"><font color="#000000">单 
                                  价</font></td>
                                <td bgcolor="#FFFFFF">
<%if rsc("ifdisc")=true then%> 
<%session("price")=FormatNumber(csng(rsc("price"))*session("y_it_userdiscount"),2)&"元/"&rsc("subsunit")%> 
<%else%> 
<%session("price")=FormatNumber(csng(rsc("price")),2)&"元/"&rsc("subsunit")%> 
<%end if%>
                                  <%=session("price")%> </td>
                              </tr>
                              <tr> 
                                <td height="25" bgcolor="#f5f5f5"><font color="#000000">数 
                                  量</font></td>
                                <td bgcolor="#FFFFFF"> <select name="<%="Q_" & rsc("bookbm")%>">
<%
for i=1 to 12 
if i=1 then
selectedstr="selected"
else
selectedstr=""
end if
%>
                                    <option <%=selectedstr%> value=<%=i%>>&nbsp; <%=i%>&nbsp;<%=rsc("subsunit")%> </option>
<%
next
%>
                                 </select>
                                 </td>
                              </tr>
                              <%
rsc.MoveNext
Wend
rsc.close
set rsc=nothing
%>
                              <tr> 
                                <td height="25" bgcolor="#f5f5f5">请选择支付方式</td>
                                <td bgcolor="#FFFFFF"> <select name="paymenttype">
                                    <option value="" selected>请选择支付方式</option>
                                    <%
sqlp="select paymenttype from paydefault"
set rsp=server.createobject("ADODB.Recordset")
rsp.open sqlp,conn,1,1
while not rsp.eof
if rsp("paymenttype")<>"预付款支付" then 
%>
                                    <option value="<%=rsp("paymenttype")%>"><%=rsp("paymenttype")%></option>
<%
elseif session("y_it_prepay")>0 and rsp("paymenttype")="预付款支付" then
%>
                                    <option value="<%=rsp("paymenttype")%>"><%=rsp("paymenttype")%></option>
<%
end if
rsp.movenext
wend
rsp.close
set rsp=nothing
%>
                                  </select> </td>
                              </tr>
                              <tr> 
                                <td height="50" valign="middle" bgcolor="#f5f5f5"><p><font color="#FF0000">特 
                                    别 说 明</font></p>
                                  </td>
                                <td height="50" valign="middle" bgcolor="#FFFFFF"><textarea name="ordernote" cols="55" rows="12" id="ordernote">绑定的域名：
( 注：如已有域名，请填上域名并在域名后加注明“已有”
  例如：“www.abc.com 已有” )

邮箱个数：

邮箱大小（每个）：
( 注：一般是10M  25M  50M每个，如特别需要请说明 )</textarea></td>
                              </tr>
                              <tr> 
                                <td height="50" colspan="2" valign="middle" bgcolor="#FFFFFF"> 
                                  <div align="center"> </div></td>
                              </tr>
                              <tr bgcolor="#f5f5f5"> 
                                <td height="35" colspan="2"><div align="center"><font color="#000000"> 
                                    <input type="submit" class=botton name="payment" value="下 一 步 >>">
                                    <input type="hidden" name="cmdShow" value="Yes">
                                    </font></div></td>
                              </tr>
                              <tr> 
                                <td height="35" colspan="2" bgcolor="#FFFFFF"><table width="88%" border="0" cellspacing="0" cellpadding="0" align="center">
                                    <tr> 
                                      <td height="150"><font color="#285278">一、请检查以上的商品清单各项内容，选取购买产品的年限<br>
                                        <br>
                                        二、选取好你的支付方式，建设您尽可能不用“邮局汇款”进行支付<br>
                                        <br>
                                        三、没有疑问的请点击“下一步”继续 ，如有疑问请<a href="emailtome.asp" target="_blank"><font color="#FF0000">点此在线提交</font></a> 
                                        <%session("subslist")= subslist%>
                                        <%session("Quatityt")= Quatityt%>
                                        <%session("gyslist")= gyslist%>
                                        </font></td>
                                    </tr>
                                  </table></td>
                              </tr>
                            </form>
                          </table>
            
            
                          <br> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td>&nbsp;</td>
                            </tr>
                          </table>
                          <table width="100%" border="0" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td>&nbsp;</td>
                            </tr>
                          </table></TD>
                      </TR>
                    </TBODY>
                  </TABLE>
                </td>
        </tr>
      </table>
</td>
          <td width="7" background="images/right-7x5.gif">&nbsp;</td>
        </tr>
      </table></td>
    <td width="10" background="images/dnyes-bg-left-and-right.gif"><img src="images/dnyes-bg-left-and-right.gif" width="10" height="1"></td>
  </tr>
</table>
<%
session("ordernote")=request("ordernote")
if request("paymenttype")<>"" then
	if request("payment")="下 一 步 >>" then
	response.redirect "payment.asp?paymenttype="&request("paymenttype")
	end if
end if
%>
<!--#include file="copyright.asp"-->
</body>
</html>
                