<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<!--#include file="css.asp"-->
<!--#include file="myjmail.asp" -->
<%on error resume next
%>
<script language="JavaScript">
function CheckForm()
{

	if (document.modiinfo.uid.value.length == 0) {
		alert("请输入您的帐号.");
		document.modiinfo.uid.focus();
		return false;
	}
	if (document.modiinfo.email.value.length == 0) {
		alert("请输入您的邮箱.");
		document.modiinfo.email.focus();
		return false;
	}
	return true;
}
</script>
<%
if session("y_it_uid")<>"" then response.Redirect("error.asp?err=013")
uid=request("uid")
email=request("email")
%>
<%
if uid<>"" then
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from user where uid='"&uid&"'"
rs.open sql,conn,1,1
%>
<%
if rs.eof or rs.bof then
%>
<script language="JavaScript">
alert("输入的帐号不存在，你重新输入")
</script>
<%
rs.close
set rs=nothing
%>
<%
else 
if trim(rs("email"))<>trim(email) then
%>
<script language="JavaScript">
alert("输入的邮箱错误，你重新输入")
</script>
<%
rs.close
set rs=nothing
%>
<%
else
topic="信网 数信科技 - 取回密码信"
passlink="<a href=http://www.dnyes.com/usermodpass11.asp?uid="&uid&"&oldpass="&rs("pwd")&" target=_blank>点击这里返取回密码</a>"
call getemailbody3(emailbody,uid,passlink)
call Jmail2(email,topic,emailbody)
response.Redirect("callmypassok.asp")
response.End()
end if
end if
end if
%>
<html>
<head>
<title><%=Application("y_it_itname")%>-取回密码</title>
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
                 <TABLE cellSpacing=0 cellPadding=0 width=100% border=0>
                    <TBODY>
                      <TR> 
                        <TD width="34%" vAlign=center style="BORDER-left: #CCCCCC 1px solid; BORDER-RIGHT: #CCCCCC 1px solid; BORDER-BOTTOM: #CCCCCC 1px solid;BORDER-TOP: #CCCCCC 1px solid"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td width="12" height="25" background="images/host-2.gif">&nbsp;</td>
                              <td valign="bottom" background="images/host-2.gif">取 回 密 码 
                                </td>
                            </tr>
                          </table>
                          <table width="75%" border="0" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td height="6"><img src="1.gif" width="1" height="1"></td>
                            </tr>
                          </table>
                          <table width="498" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                            <form name="modiinfo" method="POST" action="callmypass.asp" onSubmit="return CheckForm();">
                              <tr bgcolor="#f5f5f5"> 
                                <td height="80" colspan="2">您可以在此方便的取回您的密码<br>
                                  <br>
                                  请在下面填上你的客户帐号和邮箱<br>
                                  <br>
                                  我们会把相关资料发到你的邮箱--（邮箱不正确，不能进入） </td>
                              </tr>
                              <tr bgcolor="#FFFFFF"> 
                                <td height="60" colspan="2"><div align="center"> 
                                    <input type="hidden" name="pwdold" value="<%=pwdold%>">
                                    <input type="hidden" name="pwdcon" value=<%=rs("pwd")%>>
                                  </div></td>
                              </tr>
                              <br>
                              <br>
                              <tr> 
                                <td height="35" bgcolor="#f5f5f5">帐&nbsp;&nbsp;&nbsp;&nbsp;号： </td>
                                <td width="51%" height="33" bgcolor="#FFFFFF"> 
                                  &nbsp; <input type="text" name="uid" maxlength="20" class="form" size="20"></td>
                              </tr>
                              <tr> 
                                <td width="27%" height="35" bgcolor="#f5f5f5">邮&nbsp;&nbsp;&nbsp;&nbsp;箱：</td>
                                <td bgcolor="#FFFFFF">&nbsp; <input type="text" name="email" maxlength="40" class="form" size="20"></td>
                              </tr>
                              <tr> 
                                <td height="60" colspan="2" bgcolor="#FFFFFF">&nbsp;</td>
                              </tr>
                              <tr bgcolor="#f5f5f5"> 
                                <td height="35" colspan="2"><div align="center"> 
                                    <input class=botton type="submit" name="ok" value="确 定">
                                    &nbsp;&nbsp; 
                                    <input class=botton type="reset" name="Reset" value="清 除">
                                  </div></td>
                              </tr>
                            </form>
                          </table>
                          <br>
                          <table width="100%" border="0" cellpadding="0" cellspacing="0">
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