<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<!--#include file="css.asp"-->
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
<%
if session("y_it_uid")="" then
response.Redirect("error.asp?err=014&zurl=connect.asp")
response.End()
end if
%>
<html>
<head>
<title>����-���ſƼ� ��ϵ����</title>
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
                      <td width="82%" valign="bottom"><font color="#FFFFFF"><b>�� 
                        �� �� ¼</b></font></td>
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
                      <td width="82%" valign="bottom"><font color="#FFFFFF"><b>�� 
                        �� �� ��</b></font></td>
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
                                    <td valign="bottom" background="images/host-2.gif">�� ϵ �� ��</td>
                                  </tr>
                                </table>
                                <table width="75%" border="0" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td height="6"><img src="1.gif" width="1" height="1"></td>
                                  </tr>
                                </table>
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
mailfrom="�ʺţ�"&uid&"/��ϵ�ˣ�"&contact
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
JMail.From = email '�����˵�����
JMail.FromName = mailfrom '�����˵�����
JMail.Subject = subject '�ʼ�������
JMail.Body = mailtext '�ʼ�������
'==============================�ռ��˵ĵ�ַ��
JMail.AddRecipient "support@dnyes.com" ' �ռ��˵ĵ�ַ
'JMail.AddRecipientBcc "oudream@21cn.com"  �ռ��˵ĵ�ַ
JMail.Priority=5 '�ʼ�����1-5����Խ�󼶱�Խ��---3Ϊ��ͨ�ʼ�
JMail.Send("218.85.132.56") '��ɫ�������ʼ���������ַ
JMail.Close
Set JMail=nothing 
if err then 
err.clear
Response.Write "<center><b> ���Ź����Ѿ��򿪣������������֧�ַ��Ż��������ַ���󣬵����ż��޷�������</b></center>"
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
                              width="87" align="center" bgColor=#f5f5f5>�ͻ��ʺţ�</TD>
        <TD width="388" bgColor=#FFFFFF><%=rs("uid")%></TD>
      </TR>
      <TR bgColor=#ffffff> 
                                        <TD align="center" bgColor=#f5f5f5>�� ϵ �ˣ�</TD>
        <TD bgColor=#FFFFFF><%=rs("contactz")%></TD>
      </TR>
      <TR bgColor=#ffffff> 
                                        <TD align="center" bgColor=#f5f5f5>�����ַ��</TD>
        <TD bgColor=#FFFFFF><%=rs("email")%></TD>
      </TR>
      <TR bgColor=#ffffff> 
                                        <TD align="center" bgColor=#f5f5f5> �ʼ����⣺</TD>
        <TD bgColor=#FFFFFF><INPUT name=subject  id="subject"
                              size=48></TD>
      </TR>
      <TR bgColor=#ffffff> 
                                        <TD align="center" bgColor=#f5f5f5>�ʼ����ݣ�</TD>
                                        <TD bgColor=#FFFFFF><textarea cols=55 name=mailtext id="mailtext" rows=10></textarea> 
                                        </TD>
      </TR>
      <TR bgColor=#ffffff>
        <TD bgColor=#f5f5f5>&nbsp;</TD>
        <TD bgColor=#FFFFFF><input class=botton name=cc type=submit value=ȷ���ύ&gt;&gt;></TD>
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
                              <td><table width="498" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                                  <tr> 
                                    <td width="87" align="center" bgcolor="#f5f5f5">��˾���ƣ�</td>
                                    <td width="388" bgcolor="#FFFFFF">�㶫ʡ˳�����ſƼ����޹�˾</td>
                                  </tr>
                                  <tr> 
                                    <td align="center" bgcolor="#f5f5f5">��˾�绰��</td>
                                    <td bgcolor="#FFFFFF">0757-22205321��26823530��23808530</td>
                                  </tr>
                                  <tr> 
                                    <td align="center" bgcolor="#f5f5f5">��˾���棺</td>
                                    <td bgcolor="#FFFFFF">0757-22205321</td>
                                  </tr>
                                  <tr> 
                                    <td align="center" bgcolor="#f5f5f5">��˾��ַ��</td>
                                    <td bgcolor="#FFFFFF">�㶫ʡ˳����������ҵ���ã����ϵ»����棩��¥C3010-C3011</td>
                                  </tr>
                                  <tr> 
                                    <td align="center" bgcolor="#f5f5f5">д��¥��</td>
                                    <td bgcolor="#FFFFFF">�㶫ʡ˳������������½�16��301</td>
                                  </tr>
                                  <tr> 
                                    <td align="center" bgcolor="#f5f5f5">�������룺</td>
                                    <td bgcolor="#FFFFFF">528300</td>
                                  </tr>
                                  <tr> 
                                    <td align="center" bgcolor="#f5f5f5"> E - MAIL��</td>
                                    <td bgcolor="#FFFFFF">�ͻ�����sales@dnyes.com&nbsp;&nbsp;�ͻ�֧�֣�support@dnyes.com</td>
                                  </tr>
                                  <tr> 
                                    <td rowspan="2" align="center" bgcolor="#f5f5f5">OICQ֧�֣�</td>
                                    <td bgcolor="#FFFFFF">�ͻ�����30013002&nbsp;&nbsp;<img src="http://icon.tencent.com/30013002/l/612/" width="68" height="18"> 
                                      �ͻ�֧�֣�30013004&nbsp;&nbsp;<img src="http://icon.tencent.com/30013004/l/612/" width="68" height="18"> 
                                    </td>
                                  </tr>
                                  <tr>
                                    <td bgcolor="#FFFFFF">QQ����ʱ��Ϊ����8��30-11��30������13��30-18��00������ʱ�����õ绰��������ʼ���ϵ��</td>
                                  </tr>
                                </table></td>
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