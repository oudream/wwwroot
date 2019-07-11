<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<%on error resume next
%>
<script language="JavaScript">
function CheckForm()
{

	if (document.modiinfo.pwdold.value.length == 0) {
		alert("请输入旧密码.");
		document.modiinfo.pwdold.focus();
		return false;
	}
	if (document.modiinfo.pwd.value.length == 0) {
		alert("请输入新密码.");
		document.modiinfo.pwd.focus();
		return false;
	}
	if (document.modiinfo.pwd.value != document.modiinfo.PW_Again.value) {
		alert("您两次输入的密码不一样！请重新输入.");
		document.modiinfo.pwd.focus();
		return false;
	}
   if (document.modiinfo.pwd.value.length < 6)
  {
    alert("密码不能少于6个字符！")
    document.modiinfo.pwd.focus()
    return false
  }
  if (document.modiinfo.pwd.value.length > 20)
  {
    alert("密码不能超过20个字符！")
    document.modiinfo.pwd.focus()
    return false
  }      
	
	return true;
}
</script>
<%
if session("y_it_uid")="" then
response.redirect "error.asp?err=014&zurl=usermodpass.asp"
response.end
end if
%>
<%
if trim(request("errorstr"))="yes" then
%>
<script language="JavaScript">
alert("你输入的旧密码错误，请重新输入")
</script>
<%
nowtime=now()
end if
%>
<html>
<head>
<title><%=Application("y_it_itname")%>-修改密码</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css.css" type="text/css">
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
          <td width="229" valign="top" background="images/left-228x5.gif"> 
            <table width="229" border="0" cellpadding="0" cellspacing="0">
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
                              <td valign="bottom" background="images/host-2.gif">修 改 密 码</td>
                            </tr>
                          </table>
                          <table width="75%" border="0" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td height="6"><img src="1.gif" width="1" height="1"></td>
                            </tr>
                          </table>
                          <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td colspan="2" height="1"></td>
                            </tr>
                            <tr> 
                              <td width="1"></td>
                              <td width="564"> <table width="498" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                                  <%
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from user where uid='"&session("y_it_uid")&"'"
rs.open sql,conn,1,1
%>
                                  <form name="modiinfo" method="POST" action="usermodpassok.asp" onSubmit="return CheckForm();">
                                    <tr> 
                                      <td height="80" colspan="2" bgcolor="#f5f5f5"> 
                                        您可以在此方便的修改您的登陆密码<br>
                                        <br>
                                        如果您更改了您的密码，请在下次登陆时使用新密码 <br>
                                        <br>
                                        密码更改后系统会自动退出登陆状态，请记好您的新密码，然后重新登陆</td>
                                    </tr>
                                    <tr> 
                                      <td width="27%" height="30" bgcolor="#f5f5f5">用户帐号：</td>
                                      <td bgcolor="#FFFFFF">&nbsp; <%=rs("uid")%></td>
                                    </tr>
                                    <tr> 
                                      <td height="31" bgcolor="#f5f5f5">旧的密码：</td>
                                      <td width="51%" bgcolor="#FFFFFF">&nbsp; 
                                        <input type="password" name="pwdold" maxlength="21" class="form" size="21"> 
                                        <input type="hidden" name="pwdcon" value=<%=rs("pwd")%>> 
                                      </td>
                                    </tr>
                                    <br>
                                    <br>
                                    <tr> 
                                      <td height="28" bgcolor="#f5f5f5">新&nbsp;密&nbsp;码： </td>
                                      <td height="28" bgcolor="#FFFFFF"> &nbsp; 
                                        <input type="password" name="pwd" maxlength="21" class="form" size="21"></td>
                                    </tr>
                                    <tr> 
                                      <td width="27%" height="30" bgcolor="#f5f5f5">重复密码：</td>
                                      <td bgcolor="#FFFFFF">&nbsp; <input type="password" name="PW_Again" maxlength="21" class="form" size="21"></td>
                                    </tr>
                                    <tr bgcolor="#FFFFFF"> 
                                      <td height="50" colspan="2">&nbsp;</td>
                                    </tr>
                                    <tr bgcolor="#f5f5f5"> 
                                      <td height="35" colspan="2"><div align="center"> 
                                          <input class=botton type="submit" name="ok" value="修 改">
                                          &nbsp;&nbsp; 
                                          <input class=botton type="reset" name="Reset" value="清 除">
                                        </div></td>
                                    </tr>
                                  </form>
                                </table></td>
                            </tr>
                            <%
rs.close
set rs=nothing
%>
                            <tr> 
                              <td colspan="2" height="1"></td>
                            </tr>
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
