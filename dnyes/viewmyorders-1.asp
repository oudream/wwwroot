<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<!--#include file="css.asp"-->
<%on error resume next%>
<%
if session("y_it_uid")="" then
response.redirect "index.asp"
response.end
end if
%>
<%
id=request("id")
if id="" then
response.redirect "viewmyorders.asp"
response.end
end if
%>
<%
sql="select * from orders where id="&id
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs("username")<>session("y_it_uid") then
rs.close
set rs=nothing
response.redirect "viewmyorders.asp"
response.end
end if
sqlt="select * from user where uid='"&session("y_it_uid")&"'"
set rst=server.createobject("ADODB.Recordset")
rst.open sqlt,conn,1,1
%>
<html>
<head>
<title><%=Application("y_it_itname")%>-查看我的定单</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="inc/southdns.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
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
                <td height="180" valign="top"> 
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
                              <td valign="bottom" background="images/host-2.gif">查看定单详细情况</td>
                            </tr>
                          </table>
                          <table width="75%" border="0" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td height="6"><img src="1.gif" width="1" height="1"></td>
                            </tr>
                          </table>
                          <table width="498" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#cccccc">
                            <tr bgcolor="#f5f5f5"> 
                              <td width="28%" height="35">以下是您的定单 <%=rs("inBillNo")%> 的详细信息，请您查阅</td>
                            </tr>
                          </table>
                          <br>
                          <table width="498" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                            <tr bgcolor="#f5f5f5"> 
                              <td height="25" width="73">订 单 号：</td>
                              <td><%=rs("inBillNo")%></td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td height="25">产品名称：</td>
                              <td><%=rs("xiangxi")%></td>
                            </tr>
                            <tr bgcolor=#f5f5f5> 
                              <td height="25">服务商：</td>
                              <td><%=rs("xxgys")%></td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td height="25">订购时间：</td>
                              <td><b><%=rs("ordertime")%></b></td>
                            </tr>
                            <tr bgcolor="#FFCC33"> 
                              <td height="22" bgcolor="#FFFFFF">购买年限：</td>
                              <td width="298" bgcolor="#FFFFFF"><%=rs("xiangxi1")%>年</td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td height="5">订单金额：</td>
                              <td height="22"><%=rs("summoney")%> RMB</td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td height="5">支付方式：</td>
                              <td height="22"><%=rs("paymenttype")%></td>
                            </tr>
                            <tr bgcolor="#FFFFFF"> 
                              <td height="5">订单状态：</td>
                              <td height="22"> <%if rs("ifcomp")=true then
response.write "已经开通"
else
if rs("ifpay")=true then
response.write "付款已到，尚未开通。"
else
response.write "尚未付款"
end if
end if
%> </td>
                            </tr>
                            <%if rs("iffalse")=true then%>
                            <tr bgcolor="#FFFFFF"> 
                              <td height="5"><font color="red"><b>此订单错误信息：</b></font></td>
                              <td height="22"> <%
errmsg=replace(rs("errmsg"),chr(13),"<br>")
errmsg=replace(errmsg,chr(32),"&nbsp;")
response.write errmsg
%> </td>
                            </tr>
                            <%end if%>
                            <tr bgcolor="#FFFFFF"> 
                              <td height="5">开通时间：</td>
                              <td height="22"><%=rs("kttime")%></td>
                            </tr>
                            <tr bgcolor="#FFFFFF">
                              <td height="5">特别说明：</td>
                              <td height="22">
<%
ordernote = replace(rs("ordernote"),chr(13),"<br>")
%>	
<%=ordernote%>						  
							  </td>
                            </tr>
                          </table>
                          <br>
                          <br>
                          
<%
rs.close
set rs=nothing
rst.close
set rst=nothing
%>
                          <table width="100%" border="0" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td align="center"><br>
                                <br>
								<a href='javascript:window.close();'><font color="#330099">[关闭窗口]</font></a>								
                                <br>
                                <br>
                              </td>
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
