<script language="JavaScript">
function CheckForm()
{
	if (document.formmail.subject.value.length == 0) {
		alert("�������ʼ�����.");
		document.formmail.subject.focus();
		return false;
	}
	if (document.formmail.mailtext.value.length == 0) {
		alert("�������ʼ�����.");
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
                                    <td bgcolor="#f5f5f5">˵��:<br>
                                      1�������������﷽��Ľ��������ύ<br>
                                      2������ʱ�뾡����˵���Ա����ǵĹ�����Ա���õ����͸���Ľ����������.<br>
                                      3����Ҳ����ͨ�����ǵ�FAQ����Ѱ��������Ҫ�Ĵ�(<a href="dnyesfaq.asp"><font color="#FF0000">��˽���FAQ����ҳ��</font></a>)<br>
                                      4���յ����µ���������ǻᾡ��ĸ���ظ�,лл.</td>
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
JMail.From = zemail '�����˵�����
JMail.FromName = mailfrom '�����˵�����
JMail.Subject = subject '�ʼ�������
JMail.Body = mailtext '�ʼ�������
'==============================�ռ��˵ĵ�ַ��
JMail.AddRecipient "oudream@21cn.com" ' �ռ��˵ĵ�ַ
'JMail.AddRecipientBcc "oudream@21cn.com"  �ռ��˵ĵ�ַ
JMail.Priority=5 '�ʼ�����1-5����Խ�󼶱�Խ��---3Ϊ��ͨ�ʼ�
JMail.Send("218.85.132.56") '��ɫ�������ʼ���������ַ
JMail.Close
Set JMail=nothing 
if err then 
err.clear
Response.Write "<br><br><center><b> ���Ź����Ѿ��򿪣����������æ�������ż��޷�������<br><br>�����κ����������ţ�<a href=mailto:" & Application("y_it_fromemail") &">" & Application("y_it_fromemail")&"</a></b></center>"
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
                              width="23%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>�� 
                ϵ �ˣ�</TD>
              <TD bgColor=#ffffff width="77%"><INPUT name="mailfrom"  id="mailfrom"
                              size=50></TD>
            </TR>
            <TR bgColor=#ffffff> 
              <TD 
                              width="23%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>�����ַ��</TD>
              <TD bgColor=#ffffff width="77%"><INPUT name="zemail"  id="zemail"
                              size=50></TD>
            </TR>
            <TR bgColor=#ffffff> 
              <TD 
                              width="23%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5> 
                ��&nbsp;&nbsp;&nbsp;&nbsp;�⣺</TD>
              <TD bgColor=#ffffff width="77%"><INPUT name=subject  id="subject"
                              size=50></TD>
            </TR>
            <TR bgColor=#ffffff> 
              <TD 
                              width="23%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>�ʼ����ݣ�</TD>
              <TD bgColor=#ffffff width="77%"><textarea cols=49 name=mailtext id="mailtext" rows=10></textarea> 
              </TD>
            </TR>
            <TR bgColor=#ffffff> 
              <TD bgColor=#f5f5f5 borderColor=#c9c9c9>&nbsp;</TD>
              <TD bgColor=#ffffff><input class=botton name=cc type=submit value=�����ύ&gt;&gt;></TD>
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