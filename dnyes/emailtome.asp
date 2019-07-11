<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<script language="JavaScript">
function CheckForm()
{
	if (document.formmail.subject.value.length == 0) {
		alert("请输入邮件标题.");
		document.formmail.subject.focus();
		return false;
	}
	if (document.formmail.mailtext.value.length == 0) {
		alert("请输入邮件内容.");
		document.formmail.mailtext.focus();
		return false;
	}
		return true;
}
</script>
<%
if session("y_it_uid")="" then
response.Redirect("error.asp?err=014&zurl=emailtome.asp")
response.End()
end if
%>
<html>
<head>
<title><%=Application("y_it_itname")%>-查看资料</title>
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
                        <TD width="34%" vAlign=center style="BORDER-left: #CCCCCC 1px solid; BORDER-RIGHT: #CCCCCC 1px solid; BORDER-BOTTOM: #CCCCCC 1px solid;BORDER-TOP: #CCCCCC 1px solid"> 
                          <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td height="8" colspan="2"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td width="12" height="25" background="images/host-2.gif">&nbsp;</td>
                                    <td valign="bottom" background="images/host-2.gif">问 题 在 线 提 交</td>
                                  </tr>
                                </table>
                                <table width="75%" border="0" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td height="6"><img src="1.gif" width="1" height="1"></td>
                                  </tr>
                                </table>
                                <table width="498" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                                  <tr>
                                    <td bgcolor="#f5f5f5">说明:<br>
                                      1、您可以在这里方便的进行在线提交<br>
                                      2、提问时请尽可能说清以便我们的工作人员更好的理解和更快的解决您的疑问.<br>
                                      3、您也可以通过我们的FAQ试着寻找您所需要的答案(<a href="dnyesfaq.asp"><font color="#FF0000">点此进入FAQ帮助页面</font></a>)<br>
                                      4、收到阁下的问题后我们会尽快的给你回复,谢谢.</td>
                                  </tr>
                                </table><br>
                                <%
if session("y_it_uid")="" then
response.redirect "index.asp"
response.end
end if
%>

<%
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from user where uid='"&session("y_it_uid")&"'"
rs.open sql,conn,1,1
%>

<%
uid=rs("uid")
contact=rs("contactz")
mailfrom="帐号："&uid&"/联系人："&contact
email=rs("email")
mailtext=request("mailtext")
subject=request("subject")
%>

<%
if mailtext<>"" then

Set JMail=Server.CreateObject("JMail.Message")
JMail.Logging=True
JMail.Charset="gb2312"
JMail.ContentType = "text/html"
JMail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")
JMail.MailServerUserName = "postmaster@dnyes.com"
JMail.MailServerPassword = "winer031203"
JMail.From = email '发件人的信箱
JMail.FromName = mailfrom '发件人的名字
JMail.Subject = subject '邮件的主题
JMail.Body = mailtext '邮件的内容
'==============================收件人的地址！
JMail.AddRecipient "support@dnyes.com" ' 收件人的地址
'JMail.AddRecipientBcc "oudream@21cn.com"  收件人的地址
JMail.Priority=5 '邮件级别1-5数字越大级别越高---3为普通邮件
JMail.Send("218.85.132.56") '红色变量是邮件服务器地址
JMail.Close
Set JMail=nothing 
if err then 
err.clear
Response.Write "<br><br><center><b> 发信功能已经打开，但因服务器忙，导致信件无法发出！<br><br>如有任何疑问请至信：<a href=mailto:" & Application("y_it_fromemail") &">" & Application("y_it_fromemail")&"</a></b></center>"
response.End()
else
response.Redirect("emailtomeok.asp")
end if
end if
%>


<FORM action="emailtome.asp" method=post name="formmail" id="formmail" onSubmit="return CheckForm();">
                        
  <TABLE width="498" border=0 align="center" cellPadding=5 
                        cellSpacing=1 bgColor=#CCCCCC>
    <TBODY>
      <TR bgColor=#ffffff> 
                                        <TD 
                              width="23%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>客户帐号：</TD>
        <TD bgColor=#ffffff width="77%"><%=rs("uid")%></TD>
      </TR>
      <TR bgColor=#ffffff> 
                                        <TD 
                              width="23%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>联 系 人：</TD>
        <TD bgColor=#ffffff width="77%"><%=rs("contactz")%></TD>
      </TR>
      <TR bgColor=#ffffff> 
                                        <TD 
                              width="23%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>邮箱地址：</TD>
        <TD bgColor=#ffffff width="77%"><%=rs("email")%></TD>
      </TR>
      <TR bgColor=#ffffff> 
                                        <TD 
                              width="23%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5> 
                                          标&nbsp;&nbsp;&nbsp;&nbsp;题：</TD>
        <TD bgColor=#ffffff width="77%"><INPUT name=subject  id="subject"
                              size=50></TD>
      </TR>
      <TR bgColor=#ffffff> 
                                        <TD 
                              width="23%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>邮件内容：</TD>
                                        <TD bgColor=#ffffff width="77%"><textarea cols=49 name=mailtext id="mailtext" rows=10></textarea> 
                                        </TD>
      </TR>
      <TR bgColor=#ffffff>
        <TD bgColor=#f5f5f5 borderColor=#c9c9c9>&nbsp;</TD>
        <TD bgColor=#ffffff><input class=botton name=cc type=submit value=在线提交&gt;&gt;></TD>
      </TR>
    </TBODY>
  </TABLE>
</FORM>


<%
rs.close
set rs=nothing
%>
                              </td>
                            </tr>
                          <tr> 
                              <td width="70%">&nbsp; </td>
                            </tr>
                            <tr> 
                              <td>&nbsp;</td>
                            </tr>
                          </table>
                          <br> <br> </TD>
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