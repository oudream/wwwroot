<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<!--#include file="css.asp"-->
<%on error resume next%>
<%
id=request("id")
sql="select * from myfaq where id="&id
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
newstype=rs("newstype")
subject=rs("subject")
message=rs("message")
message=replace(message,chr(13),"<br>")
idate=rs("idate")
rs.close
set rs=nothing
%>
<html>
<head>
<title><%=Application("y_it_itname")%>-<%=newstype%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
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
                        户 服 务</b></font></td>
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
                <td width="2" valign="top" background="images/left-2-right.gif"><img src="images/left-2-right.gif" width="2" height="5"></td>
              </tr>
            </table></td>
          <td width="7">&nbsp;</td>
          <td width="513" valign="top"> <table width="500" align="center">
              <tr> 
                <td align="center">产品相关流程</td>
              </tr>
              <tr> 
                <td align="left">域名注册流程</td>
              </tr>
              <tr> 
                <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="13" rowspan="5">&nbsp;</td>
                      <td height="20" colspan="2">1.选择您需要的域名类型</td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2">2.查询你想要的域名是否可以注册</td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2">3.当提示可以注册时开始提交订单</td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2">4.填写好域名订单并提交</td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2">5.付款后请通知我们</td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td height="20" colspan="2">6.收取我们的确认信并开始使用</td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td width="199" height="20" align="right">&nbsp;</td>
                      <td width="17">&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
              <tr> 
                <td align="left">虚拟主机流程</td>
              </tr>
              <tr> 
                <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="13" rowspan="5">&nbsp;</td>
                      <td height="20" colspan="2">1.选择您需要的主机类型</td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2">2.注册或转入域名</td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2">3.如有特别要求请先联系我们</td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2">4.提交购买主机的订单</td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2">5.付款后请通知我们</td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td height="20" colspan="2">6.收取我们的确认信并开始使用</td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td width="199" height="20" align="right">&nbsp;</td>
                      <td width="17">&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td>企业邮箱流程</td>
              </tr>
              <tr> 
                <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="13" rowspan="5">&nbsp;</td>
                      <td height="20" colspan="2">1.选择您需要的邮箱类型</td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2">2.注册或转入域名,用于绑定邮箱</td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2">3.如有特别要求请先联系我们</td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2">4.提交购买邮箱的订单</td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2">5.付款后请通知我们</td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td height="20" colspan="2">6.收取我们的确认信并开始使用</td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td width="199" height="20" align="right">&nbsp;</td>
                      <td width="17">&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td>租用托管流程</td>
              </tr>
              <tr> 
                <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="13" rowspan="5">&nbsp;</td>
                      <td height="20" colspan="2">1.选择您需要的服务类型</td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2">2.注册或转入域名</td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2">3.如有特别要求请先联系我们</td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2">4.提交租用托管的订单</td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2">5.付款后请通知我们</td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td height="20" colspan="2">6.收取我们的确认信并开始使用</td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td width="199" height="20" align="right">&nbsp;</td>
                      <td width="17">&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td>网站建设流程</td>
              </tr>
              <tr> 
                <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="13" rowspan="5">&nbsp;</td>
                      <td height="20" colspan="2">1.选择您需要的建站类型</td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2">2.注册或转入域名</td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2">3.提交网站建设的订单</td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2">4.付款后请通知我们</td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2">5.收取我们的确认信</td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td height="20" colspan="2">6.把你的建站要求告诉我们</td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td height="20" colspan="2">7.我们进行网站建设并开通使用</td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td width="199" height="20" align="right">&nbsp;</td>
                      <td width="17">&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td>网站推广流程</td>
              </tr>
              <tr> 
                <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="13" rowspan="5">&nbsp;</td>
                      <td height="20" colspan="2">1.选择您需要的推广类型</td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2">2.选择你想推广的关键字</td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2">3.如有特别要求请先联系我们</td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2">4.提交网站推广的订单</td>
                    </tr>
                    <tr> 
                      <td height="20" colspan="2">5.付款后请通知我们</td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td height="20" colspan="2">6.收取我们的已付款的确认信</td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td height="20" colspan="2">7.我公司立即递交排名申请</td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td width="199" height="20" align="right">&nbsp;</td>
                      <td width="17">&nbsp;</td>
                    </tr>
                  </table></td>
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
