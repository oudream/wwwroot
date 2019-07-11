<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<%on error resume next%>
<%
errid=request("err")
if errid="" then err="001"
%>
<html>
<head>
<title><%=Application("y_it_itname")%>-申诉定单</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link rel="stylesheet" href="css.css" type="text/css">
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<!--#include file="top.asp" -->
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
          <td width="229" valign="top" background="images/left-228x5.gif"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="228" height="28" background="images/left-2.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="10%" height="18" valign="bottom">&nbsp; </td>
                      <td width="8%" valign="bottom"><img src="images/gogo.gif" width="6" height="15"></td>
                      <td width="82%" valign="bottom"><font color="#FFFFFF"><b>相 
                        关 知 识</b></font></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td> <TABLE width=100% border=0 cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
                    <TBODY>
                      <TR> 
                        <TD height="2"
					vAlign=top 
                style="BORDER-left: #000000 1px solid"><img src="1.gif" width="1" height="1"></TD>
                      </TR>
                    </TBODY>
                  </TABLE>
                  <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td height="100" valign="top"> 
                        <!--#include file="base-index.asp" -->
                      </td>
                    </tr>
                  </table>
                  <TABLE width=100% border=0 cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
                    <TBODY>
                      <TR> 
                        <TD height="1"
				vAlign=top 
                style="BORDER-left: #000000 1px solid"><img src="1.gif" width="1" height="1"></TD>
                      </TR>
                    </TBODY>
                  </TABLE></td>
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
            <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td>
                  <!--#include file="support.asp" -->
                </td>
              </tr>
            </table> </td>
          <td width="7">&nbsp;</td>
          <td width="513" valign="top"><table width="100%" align="center">
              <tr align="center"> 
                <td align="center"> <br>
                  === <b>操作提示</b> ===<br>
                  <br>
                  <br>
                </td>
              </tr>
              <tr> 
                <td> 
                  <table width="400" border="0" cellspacing="0" cellpadding="0" align="center">
                    <tr> 
                      <td width="50" valign="top"><br>
                        <br> 
                        <br> </td>
                      <td valign="top">您的操作有以下问题，请您仔细阅读下面提示： <br>
                        <br>
                        <br>
                        <%if errid="001" then%>
                        您不是从本服务器上提交的表单信息！！！<br>
                        <br>
                        请进行正常操作，否则本部保留对非正常或者破坏性<br>
                        <br>
                        操作提出进一步上诉的权利！！<br>
                        <br>
                        <br>
                        <br>
                        如有任何疑问请至信：<a href="mailto:<%=Application("y_it_fromemail")%>"><%=Application("y_it_fromemail")%></a> 
                        <br> 
                        <%end if%>
                        <%if errid="002" then%>
<script language="JavaScript">
alert("对不起!    您的帐号或者密码不正确\n\n请正确填写您的帐号密码");
history.go( -1 );
</script>
                        <%end if%>
                        <%if errid="003" then%>
<script language="JavaScript">
alert("对不起!    所查询标题/内容不存在");
history.go( -1 );
</script>
                        <%end if%>
                        <%if errid="004" then%>
                        此用户名已被注册<br>
                        <br>
                        请您选择其他的用户名<br>
                        <br>
                        感谢您使用本系统<br>
                        <br>
                        如有任何疑问请至信：<a href="mailto:<%=Application("y_it_fromemail")%>"><%=Application("y_it_fromemail")%></a> 
                        <br> 
                        <%end if%>
                        <%if errid="005" then%>
                        对不起，找回密码错误<br>
                        <br>
                        系统中并不存在此帐号<br>
                        <br>
                        感谢您使用本系统<br>
                        <br>
                        如有任何疑问请至信：<a href="mailto:<%=Application("y_it_fromemail")%>"><%=Application("y_it_fromemail")%></a> 
                        <br>
                        <%end if%>
                        <%if errid="006" then%>
                        <font color=red>您尚未登录</font>，请从此入口登录，或<a href="userregprotocal.asp">点此注册</a><br>
                        <br> <form action="login.asp" method="post">
                          您的帐号:
                          <input type="text" name="uid" size="16" maxlength="20" class="form">
                          &nbsp;
                          <input type="submit" name="Submit" class=botton value="登录">
                          <input type="hidden" name="url" value="add.asp?add=<%=request("add")%>">
                          <br>
                          您的密码:
                          <input type="password" name="pwd" size="16" maxlength="20" class="form">
                          &nbsp;
                          <input class=botton type="submit" name="Submit2"  value="取消">
                        </form>
                        <br>
                        如有任何疑问请至信：<a href="mailto:<%=Application("y_it_fromemail")%>"><%=Application("y_it_fromemail")%></a> 
                        <br>
                        <%end if%>
                        <%if errid="007" then%>
                        对不起！您没有正常购买或您已经确认购买，想清楚了解<br>
						请查看您的订单再重新购物<br>
                        <br>
						购物时请尊从我们<a href="servicefollow.asp">购物流程</a><br><br>
                        感谢您对我们的支持<br>
                        <br>
                        <br>
                        <br>
                        如有任何疑问请至信：<a href="mailto:<%=Application("y_it_fromemail")%>"><%=Application("y_it_fromemail")%></a> 
                        <br> 
                        <%end if%>
                        <%if errid="008" then%>
                        对不起，您的购物总金额过低不能用预付款购物。<br>
                        <br>
						请用其它方式购物，感谢您对我们的支持！
                        <br>
                        <br>
                        如有任何疑问请至信：<a href="mailto:<%=Application("y_it_fromemail")%>"><%=Application("y_it_fromemail")%></a> 
                        <br> 
                        <%end if%>
                        <%if errid="009" then%>
                        对不起，您的定单已经被申诉过了，请不要重复申诉<br>
                        <br>
                        感谢您对我们的支持！<br>
                        <br>
                        如有任何疑问请至信：<a href="mailto:<%=Application("y_it_fromemail")%>"><%=Application("y_it_fromemail")%></a>，或QQ：30013002,30013004 
                        。 <br> 
                        <%end if%>
                        <%if errid="010" then%>
                        对不起，并无此定单，请确认<br>
                        <br>
                        请您正确使用本系统<br>
                        <br>
                        感谢您对我们的支持！<br>
                        <br>
                        如有任何疑问请至信：<a href="mailto:<%=Application("y_it_fromemail")%>"><%=Application("y_it_fromemail")%></a>，或QQ：30013002,30013004 
                        。 <br> 
                        <%end if%>
                        <%if errid="011" then%>
                        对不起，<font color=red>您尚未登录</font>，请从此入口登录，或<a href="reg.asp">点此注册</a><br>
                        <br> <form action="login.asp" method="post">
                          您的帐号:
                          <input type="text" name="uid" size="16" maxlength="20" class="form">
                          &nbsp;
                          <input type="submit" name="Submit" class=botton value="登录">
                          <input type="hidden" name="url" value="index.asp">
                          <br>
                          您的密码:
                          <input type="password" name="pwd" size="16" maxlength="20" class="form">
                          &nbsp;
                          <input type="submit" name="Submit2" class=botton value="取消">
                        </form>
                        <br>
                        如有任何疑问请至信：<a href="mailto:<%=Application("y_it_fromemail")%>"><%=Application("y_it_fromemail")%></a> 
                        <br> 
                        <%end if%>
                        <br>
                        <%if errid="012" then%>
                        对不起，<font color=red>您尚未登录</font>，请从此入口登录，或<a href="userregprotocal.asp">点此注册</a><br>
                        <br>
                        <br> <form action="login.asp" method="post">
                          您的帐号:
                          <input type="text" name="uid" size="16" maxlength="20" class="form">
                          &nbsp;
                          <input type="submit" name="Submit" class=botton value="登录">
                          <input type="hidden" name="url" value="<%=session("domainurl")%>">
                          <br>
                          您的密码:
                          <input type="password" name="pwd" size="16" maxlength="20" class="form">
                          &nbsp;
                          <input type="submit" name="Submit2" class=botton value="取消">
                        </form>
                        <br>
                        <br>
                        <br>
                        如有任何疑问请至信：<a href="mailto:<%=Application("y_it_fromemail")%>"><%=Application("y_it_fromemail")%></a> 
                        <br> 
                        <%end if%>
						<br>
                        <%if errid="013" then%>
                        对不起，<font color="red">您已登录</font><br>
                        <br>
                        您已是注册用户,如想再注册或取回密码,请先注销<br>
                        <br>
                        感谢您对我们的支持！<br>
                        <br>
                        如有任何疑问请至信：<a href="mailto:<%=Application("y_it_fromemail")%>"><%=Application("y_it_fromemail")%></a> 
                        <br> 
                        <%end if%>
                        <br>
                        <%if errid="014" then%>
                        对不起，<font color=red>您尚未登录</font>，请从此入口登录，或<a href="userregprotocal.asp">点此注册</a><br>
                        <br>
                        <br> <form action="login.asp" method="post">
                          您的帐号:
                          <input type="text" name="uid" size="16" maxlength="20" class="form">
                          &nbsp;
                          <input type="submit" name="Submit" class=botton value="登录">
                          <input type="hidden" name="url" value="<%=request("zurl")%>">
                          <br>
                          您的密码:
                          <input type="password" name="pwd" size="16" maxlength="20" class="form">
                          &nbsp;
                          <input type="submit" name="Submit2" class=botton value="取消">
                        </form>
                        <br>
                        <br>
                        <br>
                        如有任何疑问请至信：<a href="mailto:<%=Application("y_it_fromemail")%>"><%=Application("y_it_fromemail")%></a> 
                        <br> 
                        <%end if%>
                        <br> 
						</td>
                    </tr>
                  </table> </td>
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
