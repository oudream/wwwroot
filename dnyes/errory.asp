<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<%on error resume next%>
<%error=request("error")%>
<html>
<head>
<title><%=Application("y_eshop_esname")%>-操作提示</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="inc/southdns.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<!--#include file="top.asp"--> 
<table width="776" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr valign="top"> 
    <td width="146"><img src="image/in54_r23_c2.gif" width="143" height="31"> 
      <table border="0" cellspacing="1" cellpadding="0" align="center" width="100%" bgcolor="#CCCCCC">
        <tr bgcolor="#FFFFFF"> 
          <td> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <form action="userlogin.asp" method="post" name="login">
                <tr> 
                  <td colspan="2" height="20"></td>
                </tr>
                <tr> 
                  <td height="22" width="35%"> 
                    <div align="right"><img src="image/in54_r26_c7.gif" width="39" height="13"></div>
                  </td>
                  <td width="65%"> 
                    <div align="center"> 
                      <input maxlength=30 name=username size=10 class="input">
                    </div>
                  </td>
                </tr>
                <tr> 
                  <td height="22"> 
                    <div align="right"><img src="image/in54_r28_c8.gif" width="27" height="12"></div>
                  </td>
                  <td> 
                    <div align="center"> 
                      <input maxlength=16 name=password size=10 type=password class="input">
                    </div>
                  </td>
                </tr>
                <tr> 
                  <td height="30" colspan="2"> 
                    <div align="center"> <a href="reg.asp"><img src="image/in54_r32_c9.gif" width="47" height="20" border="0"></a>&nbsp;&nbsp; 
                      <input type="image" border="0" name="search2" src="image/in54_r32_c13.gif">
                      <input type="hidden" name="userlogin" value="True">
                    </div>
                  </td>
                </tr>
                <tr> 
                  <td height="30" colspan="2"> 
                    <div align="center"><a href="findpassword.asp"><img src="image/in54_r35_c10.gif" width="92" height="18" border="0"></a></div>
                  </td>
                </tr>
                <tr> 
                  <td height="45" colspan="2"> 
                    <div align="center"><img src="image/in54_r40_c10.gif" width="61" height="26"></div>
                  </td>
                </tr>
              </form>
            </table>
          </td>
        </tr>
      </table>
      <img src="image/in54_r43_c2.gif" width="146" height="25"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <%
sqls="select sarea,photo from sarea order by id" 
Set rss=Server.CreateObject("ADODB.RecordSet") 
rss.Open sqls,conn,1,1
n=0
do while not rss.eof
photo=rss("photo")
photo=photo+".gif"
%> 
        <tr> 
          <td height="40"> 
            <div align="center"> <a href="showsarea.asp?sarea=<%=rss("sarea")%>" title="<%=rss("sarea")%>" target="_blank"> 
              <img src="photo/<%=photo%>" width="136" height="34" border="0"></a></div>
          </td>
        </tr>
        <%
rss.movenext
n=n+1
if n>=7 then exit do
loop
rss.close
set rss=nothing
%> 
        <tr> 
          <td height="25"> 
            <div align="right"><img src="image/in54_r77_c14.gif" width="40" height="13"></div>
          </td>
        </tr>
      </table>
      <br>
    </td>
    <td width="630"> <br>
      <table border="0" cellspacing="1" cellpadding="0" align="center" width="600" bgcolor="#5C9DBE">
              <form action="userlogin.asp" method="post" name="login">
        <tr bgcolor="#FFFFFF"> 
          <td>
            <table width="580" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr><td height="30" colspan="3">&nbsp;</td></tr>
              <tr> 
                <td width="20" height="22">&nbsp;</td>
                  <td bgcolor="#5C9DBE">&nbsp;<b>对不起，您的操作不正确，请先下面提示</b></td>
                <td width="160">&nbsp;</td>
              </tr>
              <tr><td height="30" colspan="3">&nbsp;</td></tr>
            </table>

            <table width="580" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr>
                <td width="150" valign="top"><img src="image/54-zc1.gif" width="150" height="182"></td>
                <td valign="top" width="350">
      <table border="0" cellspacing="1" cellpadding="0" align="center" width="330" bgcolor="#5C9DBE">
                    <tr bgcolor="#FFFFFF"> 
                        <td> <br>
                          <table width="300" border="0" cellspacing="0" cellpadding="0" align="center">
                            <tr>
                              <td><font color="#CC3333"><%if error="001" then%> 
                                <br>
                                对不起，您不是从本服务器上提交的表单信息！！！<br>
                                <br>
                                请进行正常操作，否则本部保留对非正常或者破坏性<br>
                                <br>
                                操作提出进一步上诉的权利！！<br>
                                <%end if%> <%if error="002" then%> <br>
                                对不起，您所选择的用户名已经被注册了.<br>
                                <br>
                                您可以点此<a href="reg.asp">重新注册</a>您的帐号.<br>
                                <br>
                                谢谢您选择我们的服务！<br>
                                <%end if%> <%if error="003" then%> <br>
                                对不起，并没有您所填写的用户帐号.<br>
                                <br>
                                您可以点此<a href="index.asp">重新登陆</a>您的帐号.<br>
                                <br>
                                或者点此<a href="reg.asp">立即注册</a>您的帐号.<br>
                                <br>
                                谢谢您选择我们的服务！<br>
                                <%end if%> <%if error="004" then%> <br>
                                对不起，您错误的填写了帐号的密码.<br>
                                <br>
                                您可以点此<a href="index.asp">重新登陆</a>您的帐号.<br>
                                <br>
                                或者点此<a href="findpassword.asp">迅速找回</a>您的帐号.<br>
                                <br>
                                谢谢您选择我们的服务！<br>
                                <%end if%> <%if error="005" then%> <br>
                                对不起，您的帐号已经登陆过了.<br>
                                <br>
                                请您不要重复登陆.点此<a href="index.asp">返回首页</a><br>
                                <br>
                                谢谢您选择我们的服务！<br>
                                <%end if%> <%if error="006" then%> 
                                <p><br>
                                  亲爱的用户，您没有注册或者没有登陆！<br>
                                  <br>
                                  请登陆，或注册。<br>
                                  <br>
                                  <a href="findpassword.asp">忘记密码？</a> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="reg.asp">我要注册 
                                  </a> <br>
                                  您的帐号： 
                                  <input type="text" name="username" size="12" maxlength="15" class="form">
                                  &nbsp;&nbsp; 
                                  <input type="submit" name="Submit" value="登陆">
                                  <input type="hidden" name="userlogin" value="True">
                                  <br>
                                  您的密码： 
                                  <input type="password" name="password" size="12" maxlength="15" class="form">
                                  &nbsp;&nbsp; 
                                  <input type="submit" name="Submit2" value="取消">
                                  <br>
                                  <%end if%> <%if error="007" then%> <br>
                                  对不起，您的购物车为空.<br>
                                  <br>
                                  点此<a href="index.asp">重新购物.</a><br>
                                  <br>
                                  谢谢您选择我们的服务！<br>
                                  <%end if%> <%if error="008" then%> <br>
                                  对不起，并没有这个帐号.<br>
                                  <br>
                                  请您检查后重新输入.<br>
                                  <br>
                                  点此<a href="findpassword.asp">重新查询.</a><br>
                                  <br>
                                  如果您还有任何疑问，请至信: <a href="mailto:<%=Application("y_eshop_fromemail")%>"><%=Application("y_eshop_fromemail")%></a> 
                                  <%end if%> <%if error="009" then%> <br>
                                  对不起，您的订单已经申诉过了.<br>
                                  <br>
                                  我们会在最短时间内处理您的订单.<br>
                                  <br>
                                  点此<a href="index.asp">返回首页.</a><br>
                                  <br>
                                  如果您还有任何疑问，请至信: <a href="mailto:<%=Application("y_eshop_fromemail")%>"><%=Application("y_eshop_fromemail")%></a> 
                                  <%end if%> <%if error="010" then%> <br>
                                  对不起，此订单并不存在.<br>
                                  <br>
                                  请您核实您的帐号和订单号后再进行申诉.<br>
                                  <br>
                                  点此<a href="usererror.asp">重新进行申诉.</a><br>
                                  <br>
                                  如果您还有任何疑问，请至信: <a href="mailto:<%=Application("y_eshop_fromemail")%>"><%=Application("y_eshop_fromemail")%></a> 
                                  <%end if%> <%if error="011" then%> <br>
                                  对不起，您并没有以管理员登陆.<br>
                                  <br>
                                  点此<a href="admin/adminlogin.asp">重新登陆.</a><br>
                                  <br>
                                  如果您还有任何疑问，请至信: <a href="mailto:<%=Application("y_eshop_fromemail")%>"><%=Application("y_eshop_fromemail")%></a> 
                                  <%end if%> <%if error="013" then%> <br>
                                  对不起，您一次的购物金额不得底于50RMB.<br>
                                  <br>
                                  点此<a href="index.asp">继续购物.</a><br>
                                  <br>
                                  如果您还有任何疑问，请至信: <a href="mailto:<%=Application("y_eshop_fromemail")%>"><%=Application("y_eshop_fromemail")%></a> 
                                  <%end if%> <%if error="016" then%> <br>
                                  对不起，您没有选择您要评论的商品.<br>
                                  <br>
                                  点此<a href="index.asp">返回首页.</a><br>
                                  <br>
                                  如果您还有任何疑问，请至信: <a href="mailto:<%=Application("y_eshop_fromemail")%>"><%=Application("y_eshop_fromemail")%></a> 
                                  <%end if%> 
                                </b></font></td>
                            </tr>
                          </table>
                          <br>
                        </td>
                    </tr>
                  </table>
                </td>
                <td width="80" valign="top">&nbsp;</td>
              </tr>
              <tr>
                  <td height="120">&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
            </table>
          </td>
        </tr></form>
      </table>
    </td>
  </tr>
</table>
<!--#include file="copyright.asp"--> 
</body>
</html>
