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
          <td width="513" valign="top"><table width="100%" border="0" cellspacing="2" cellpadding="2">
              <tr> 
                <td width="3%"><font color="#003366">■</font></td>
                <td width="92%">成为数信科技代理的优势</td>
                <td width="5%">&nbsp;</td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td width="25">→</td>
                      <td width="745"><font color="#0000FF">1、 全线产品，大小生意都可做</font><br>
                        数信科技的产品超过80个品种，最小的60多元，大的上万元。成为代理，就可以面对客户没有顾虑，不论大小订单都可以成交，不必东问西问。生意自然好做很多。<br> 
                      </td>
                    </tr>
                    <tr> 
                      <td>→</td>
                      <td><font color="#0000FF">2、 丰厚利润，回报迅速</font><br>
                        数信科技的产品代理利润高，很多产品利润高达35%，代理当然是卖一单赚一单，实实在在的进账。<br> </td>
                    </tr>
                    <tr> 
                      <td>→</td>
                      <td><font color="#0000FF">3、 全面服务，实力保证</font><br>
                        数信科技的技术专家为客户提供全天侯的服务，节假日、国庆、春节也不例外，不会出现小公司人手不足，设备不够的局面。由于网站业务是越做越多，若因上游供应商服务不到位而失去市场，被顾客要求赔款，做代理的就损失太大了。<br> 
                      </td>
                    </tr>
                    <tr> 
                      <td>→</td>
                      <td><font color="#0000FF">4、 继承优势，面对市场不手软</font><br>
                        数信科技公司是个市场导向型公司，数信科技对市场的策略一直是以强悍而闻名，但决不蛮干。所以，各位成为代理，就可以面对竞争对手毫不畏惧。我们的优势就是大家的优势，我们成长，代理就一定会成长。</td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                    <tr> 
                      <td colspan="2">&nbsp;</td>
                    </tr>
                  </table></td>
                <td>&nbsp;</td>
              </tr>
            </table>
            <table width="100%" border="0" cellspacing="2" cellpadding="2">
              <tr> 
                <td width="3%"><font color="#003366">■</font></td>
                <td width="92%">代理的条件</td>
                <td width="5%">&nbsp;</td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
                <td><TABLE border=0 cellPadding=2 cellSpacing=0 width="100%">
                    <TBODY>
                      <TR> 
                        <TD width=25>▲</TD>
                        <TD width="737">申请个人代理的条件</TD>
                      </TR>
                      <TR> 
                        <TD width=25>&nbsp;</TD>
                        <TD>&nbsp;</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>→ </TD>
                        <TD>可以独立承担民事责任的个人。</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>→ </TD>
                        <TD>可以独立为您所发展的用户开据合法有效的服务票据。</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>→ </TD>
                        <TD>可以为用户提供必要的技术服务与咨询服务。</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>→ </TD>
                        <TD>应当拥有固定的服务场所。</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>→ </TD>
                        <TD>具有比较丰富的互联网络技术经验与从业背景。</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>→ </TD>
                        <TD>具有便利的上网通讯条件及必要的设备。</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>→ </TD>
                        <TD>个人需要提交个人身份证复印件。</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>→ </TD>
                        <TD>愿意交纳个人代理预付款500元RMB以保证业务的顺利开展。</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>→ </TD>
                        <TD>愿意接受数信科技对其的审核、监督和指导。</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>→ </TD>
                        <TD>接受数信科技的"代理体系与管理规范"。</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>&nbsp;</TD>
                        <TD>&nbsp;</TD>
                      </TR>
                      <TR> 
                        <TD>▲</TD>
                        <TD>申请企业代理的条件</TD>
                      </TR>
                      <TR> 
                        <TD>&nbsp;</TD>
                        <TD>&nbsp;</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>→ </TD>
                        <TD>可以独立承担民事责任的企业单位。</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>→ </TD>
                        <TD>可以独立为您所发展的用户开据合法有效的服务票据。</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>→ </TD>
                        <TD>可以为用户提供必要的技术服务与咨询服务。</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>→ </TD>
                        <TD>应当拥有固定的办公地点。</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>→ </TD>
                        <TD>具有比较丰富的互联网络技术经验与从业背景。</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>→ </TD>
                        <TD>具有便利的上网通讯条件及必要的设备。</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>→ </TD>
                        <TD>企业需要提交企业营业执照复印件。</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>→ </TD>
                        <TD>在本地的行业内拥有一定的知名度。</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>→ </TD>
                        <TD>在本地拥有完善的销售网络与渠道。</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>→ </TD>
                        <TD>具有良好的客户服务能力与经验。</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>→ </TD>
                        <TD>愿意交纳数信科技代理预付款500元RMB以保证业务的顺利开展。</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>→ </TD>
                        <TD>愿意接受数信科技对其的审核、监督和指导。</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>→ </TD>
                        <TD>接受数信科技"代理体系与管理规范"。</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>→ </TD>
                        <TD>对数信科技代理商的定级，不排除对某些特殊贡献的代理商执行软性考核的</TD>
                      </TR>
                      <TR> 
                        <TD align=middle vAlign=top>&nbsp;</TD>
                        <TD>标准。如对数信科技的发展和市场推广做出特殊贡献或给予重要支持的代</TD>
                      </TR>
                      <TR> 
                        <TD align=middle vAlign=top>&nbsp;</TD>
                        <TD>理商，经数信科技确认后，可酌情提供相应的优待和支持。</TD>
                      </TR>
                    </TBODY>
                  </TABLE></td>
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td><font color="#003366">■</font></td>
                <td>代理步骤与流程</td>
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td width="25">→</td>
                      <td width="745">了解要求，注册成为用户 </td>
                    </tr>
                    <tr> 
                      <td>→</td>
                      <td>系统自动认证为正式用户</td>
                    </tr>
                    <tr> 
                      <td>→</td>
                      <td><a href="mydocument/agent002.doc"><font color="#FF0000">下载代理合同</font></a>，签字并邮寄到我公司</td>
                    </tr>
                    <tr> 
                      <td>→</td>
                      <td>将预付款汇入指定帐户</td>
                    </tr>
                    <tr> 
                      <td>→</td>
                      <td>相关文件传真或扫描至我司</td>
                    </tr>
                    <tr> 
                      <td>→</td>
                      <td>域名网渠道管理正式将申请者从用户升级为代理</td>
                    </tr>
                  </table>
                </td>
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td><font color="#003366">■</font></td>
                <td>代理审核标准</td>
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
                <td><P>数信科技代理商主要有两大类个人代理商和机构代理商， 具体代理级别如下：<BR>
                    1. 铜牌代理商<BR>
                    2. 银牌代理商<BR>
                    3. 金牌代理商<BR>
                  </P>
                  <P><BR>
                    新加入的标准</P>
                  <TABLE border=0 cellPadding=0 cellSpacing=1 
                  width="85%">
                    <TBODY>
                      <TR> 
                        <TD width="35%" height=20>代理级别</TD>
                        <TD width="65%" height=20>代理条件</TD>
                      </TR>
                      <TR> 
                        <TD height=20 bgcolor="#ffffff">铜牌代理商</TD>
                        <TD height=20 bgcolor="#ffffff">预收款≥500元人民币</TD>
                      </TR>
                      <TR> 
                        <TD height=20 bgcolor="#ffffff">银牌代理商</TD>
                        <TD height=20 bgcolor="#ffffff">预收款≥3000元人民币</TD>
                      </TR>
                      <TR> 
                        <TD height=20 bgcolor="#ffffff">金牌代理商</TD>
                        <TD height=20 bgcolor="#ffffff">预收款≥10000元人民币</TD>
                      </TR>
                    </TBODY>
                  </TABLE>
                  <br>
                  <br>
                  <a href="mydocument/agent002.doc"><font color="#FF0000">点此下载代理合同</font></a> 
                </td>
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
            </table> </td>
          <td width="7" background="images/right-7x5.gif">&nbsp;</td>
        </tr>
      </table></td>
    <td width="10" background="images/dnyes-bg-left-and-right.gif"><img src="images/dnyes-bg-left-and-right.gif" width="10" height="1"></td>
  </tr>
</table>
<!--#include file="copyright.asp"-->
</body>
</html>
