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
<html>
<head>
<title>Email</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td height="8" colspan="2">
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
mailfrom=request("mailfrom")
zemail=request("zemail")
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
JMail.From = zemail '发件人的信箱
JMail.FromName = mailfrom '发件人的名字
JMail.Subject = subject '邮件的主题
JMail.Body = mailtext '邮件的内容
'==============================收件人的地址！
JMail.AddRecipient "oudream@21cn.com" ' 收件人的地址
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
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('ok!');</SCRIPT>")
end if
end if
%>


<FORM action="000.asp" method=post name="formmail" id="formmail" onSubmit="return CheckForm();">
                        
        <TABLE width="498" border=0 align="center" cellPadding=5 
                        cellSpacing=1 bgColor=#CCCCCC>
          <TBODY>
            <TR bgColor=#ffffff> 
              <TD 
                              width="23%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>联 
                系 人：</TD>
              <TD bgColor=#ffffff width="77%"><INPUT name="mailfrom"  id="mailfrom"
                              size=50></TD>
            </TR>
            <TR bgColor=#ffffff> 
              <TD 
                              width="23%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>邮箱地址：</TD>
              <TD bgColor=#ffffff width="77%"><INPUT name="zemail"  id="zemail"
                              size=50></TD>
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
                              </td>
                            </tr>
                          <tr> 
                              <td width="70%">&nbsp; </td>
                            </tr>
                            <tr> 
                              <td>&nbsp;</td>
                            </tr>
                          </table>
						  
						  <%=mailfrom%><br>
						  <%=zemail%><br>
						  <%=mailtext%><br>
						  <%=subject%><br>
</body>
</html>