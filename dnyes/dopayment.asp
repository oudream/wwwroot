<script language="JavaScript">
function CheckForm()
{
	if (document.modiinfo.uidcon.value.length == 0) {
		alert("请输入帐号");
		document.modiinfo.uidcon.focus();
		return false;
	}
	if (document.modiinfo.pwdcon.value.length == 0) {
		alert("请输入密码");
		document.modiinfo.pwdcon.focus();
		return false;
	}
}
</script>
<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<!--#include file="mdf.asp" -->
<!--#include file="css.asp"-->
<%
sum=session("sum")
'是否用预付款支付
ifpay=""

'产生订单号（内部和外部）及定单产生日期及其他信息
'交易日期，格式：YYYYMMDD
yy=year(date)
mm=right("00"&month(date),2)
dd=right("00"&day(date),2)
riqi=yy & mm & dd

'生成订单号所有所需元素,格式为：小时，分钟，秒
xiaoshi=right("00"&hour(time),2)
fenzhong=right("00"&minute(time),2)
miao=right("00"&second(time),2)
shijian=xiaoshi & fenzhong & miao
'生成订单号
inBillNo=riqi & "-" & shijian

username=session("y_it_uid")

paymenttype1=request("paymenttype")
if paymenttype1="预付款支付" then
sqlp1="select prepay from user where uid='"&session("y_it_uid")&"'"
set rsp1=server.createobject("adodb.recordset")
rsp1.open sqlp1,conn,1,1
if rsp1("prepay")<sum/2*2 then
response.write"对不起，您无预付款或者您的预付款不足以支付本次金额！"
rsp1.close
set rsp1=nothing
response.End()
end if
end if
%>
<%
'判断是否正常购买，如果不是，需要登陆
if session("y_it_uid")="" then 
response.redirect "error.asp?err=011"
response.end
end if
%>
<%
uidcon=request("uidcon")
pwdcon=request("pwdcon")
i=request("i")
paymenttype=request("paymenttype")
if pwdcon<>"" then
pwdcon=md5(pwdcon)
if trim(session("y_it_pwd"))<>trim(pwdcon) or trim(session("y_it_uid"))<>trim(uidcon) then
%>
<script language="JavaScript">
alert('你输入的帐号或密码不正确，请重新输入');
</script>
<%
nowtime=now()
else
response.Redirect("saveok.asp?paymenttype="&paymenttype&"&i="&i)
response.End()
end if
end if
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
<html>
<head>
<title><%=Application("y_it_itname")%>-进行支付</title>
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
          <td width="229" valign="top" background="images/left-228x5.gif"> <TABLE width=100% border=0 cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
              <TBODY>
                <TR> 
                  <TD height="10" 
                vAlign=top 
                style="BORDER-left: #000000 1px solid"><img src="1.gif" width="1" height="1"></TD>
                </TR>
              </TBODY>
            </TABLE>
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
                <td colspan="5" height="88">
<%
'判断购物车是否为空
ProductList = Session("ProductList")
if productlist<>"" then
sql="select * from subs where bookbm in ("&productlist&") order by bookbm"
Set rs = conn.Execute( sql )
else
response.redirect "error.asp?err=007"
response.end
end if
%>                  <TABLE cellSpacing=0 cellPadding=0 width=100% border=0>
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
  <tr bgcolor="#f5f5f5"> 
                              <td height="35" colspan="2"><%=zcontactz%> 以下是您个人信息，请查看。如需改动请<a href="usermod.asp" target="_blank"><font color="#FF0000">点此</font></a>进行修改</td>
  </tr>
  <tr> 
                              <td width="28%" height="25" bgcolor="#f5f5f5">E - 
                                MAIL：</td>
    <td width="72%" bgcolor="#FFFFFF"><%=zemail%></td>
  </tr>
  <tr> 
    <td height="25" bgcolor="#f5f5f5">联系电话：</td>
    <td bgcolor="#FFFFFF"><%=ztel%></td>
  </tr>
  <tr> 
    <td height="25" bgcolor="#f5f5f5">联系地址：</td>
    <td bgcolor="#FFFFFF"><%=zdizhiz%></td>
  </tr>
  <tr> 
    <td height="5" bgcolor="#f5f5f5">&nbsp;</td>
    <td bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
</table>
                          <br>
<form name=modiinfo id="modiinfo" method="POST" action= "dopayment.asp" onSubmit="return CheckForm();">
  <table width="498" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
    <%
While Not rs.EOF
%>
    <%
if instr(1,rs("subsname"),".",1)<>0 then
temp=Session("domain")&rs("subsname")
else
temp=rs("subsname")
end if
%>
    <%
if instr(1,rs("subsname"),".",1)<>0 then
admintemp=Session("domain")&rs("subsname")
else
admintemp=session("y_it_uid")
end if

i=i+1
glz="glzh"&i
glm="glma"&i
%>
    <tr bgcolor="#f5f5f5"> 
      <td height="35" colspan="2">所选购的商品：</td>
    </tr>
    <tr> 
      <td width="28%" height="25" bgcolor="#f5f5f5"> 商品名称</td>
      <td width="72%" height="20" bgcolor="#FFFFFF"><%=temp%> <div align="center"></div>
        <div align="center"></div></td>
    </tr>
    <tr> 
      <td height="20" bgcolor="#f5f5f5"> 供 应 商</td>
      <td bgcolor="#FFFFFF"><%=rs("gys")%></td>
    </tr>
    <%
rs.MoveNext
Wend
rs.close
set rs=nothing
%>
    <tr> 
      <td height="25" bgcolor="#f5f5f5">订购年限</td>
      <td bgcolor="#FFFFFF"><%=session("Quatityt")%></td>
    </tr>
    <tr> 
      <td height="25" bgcolor="#f5f5f5">总金额</td>
      <td bgcolor="#FFFFFF"><%=session("sum")%></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="25" colspan="2">&nbsp;</td>
    </tr>
    <tr> 
      <td height="25" bgcolor="#f5f5f5">确定帐号</td>
      <td width="350" bgcolor="#FFFFFF"><input type="textarea" name="uidcon" size="20" class="form" value="<%=session("y_it_uid")%>"> 
        <input type="hidden" name="<%=glz%>" size="20" class="form"></td>
    </tr>
    <tr> 
      <td height="25" bgcolor="#f5f5f5">确定密码</td>
      <td height="28" bgcolor="#FFFFFF"><input type="password" name="pwdcon" size="20" class="form"> 
        <input type="hidden" name="<%=glm%>" size="20" class="form"> </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="50" colspan="2">&nbsp;</td>
    </tr>
    <tr bgcolor="#f5f5f5"> 
      <td height="35" colspan="2"><div align="center"> 
          <input onClick=history.go(-1) type=button class=botton value="<<上一步" name=button>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    
          <input type="submit" class=botton name="ok" value="确定购买">
          <input type=hidden  value=<%=i%> name=i>
          <input type=hidden  value=<%=paymenttype%> name=paymenttype>
        </div></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="88" colspan="2"><div align="center"> </div></td>
    </tr>
  </table>
</form>                          <table width="100%" border="0" cellpadding="0" cellspacing="0">
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
                  </TABLE></td>
              </tr>
            </table></td>
          <td width="7" background="images/right-7x5.gif">&nbsp;</td>
        </tr>
      </table></td>
    <td width="10" background="images/dnyes-bg-left-and-right.gif"><img src="images/dnyes-bg-left-and-right.gif" width="10" height="1"></td>
  </tr>
</table>
<!--#include file="copyright.asp"-->
</body>
</html>