<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<!--#include file="css.asp"-->
<%
paymenttype=request("paymenttype")
saveok1=request("saveok1")
saveok2=request("saveok2")
%>
<html>
<head>
<title>���� ���ſƼ� - �ͻ������ɹ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<%
if saveok1<>"" and saveok2<>"" and paymenttype <> "" then
%>
<%
sqlp1="select * from paydefault where paymenttype='"&paymenttype&"'"
set rsp1=server.createobject("adodb.recordset") 
rsp1.open sqlp1,conn,1,1 
paymentmessage = rsp1("paymentmessage")
paymentmessage = replace(paymentmessage,chr(13),"<br>")
paymentmessage = replace(paymentmessage,chr(32),"&nbsp;")  
rsp1.close
set rsp1=nothing
%>
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
                        <TD width="34%" vAlign=center style="BORDER-left: #CCCCCC 1px solid; BORDER-RIGHT: #CCCCCC 1px solid; BORDER-BOTTOM: #CCCCCC 1px solid;BORDER-TOP: #CCCCCC 1px solid"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td width="12" height="25" background="images/host-2.gif">&nbsp;</td>
                              <td valign="bottom" background="images/host-2.gif">�� 
                                �� �� �� &gt;&gt; �� �� �� ��</td>
                            </tr>
                          </table>
                          <table width="75%" border="0" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td height="6"><img src="1.gif" width="1" height="1"></td>
                            </tr>
                          </table>
                          <table border=0 cellspacing=1 cellpadding=5 align=center width=498 bgcolor=#CCCCCC>
                            <tr bgcolor=#BFCEDD>
                              <td height=35 colspan=3 bgcolor="#f5f5f5"> ��л��ѡ�����ǵķ��� 
                              </td>
                            </tr>
                            <tr bgcolor=#FFFFFF> 
                              <td width=65% height=22 bgcolor="#FFFFFF"><p><br>
                                  ��ѡ���֧����ʽ�� ��<%=paymenttype%>֧�� </p>
                                <p><%=paymentmessage%><br>
                                  <br>
                                  <br>
                                  ���ν��׵Ķ�����Ϊ��<%=saveok1%> <br>
                                  <br>
                                  <br>
                                  ���ν��׵��ܽ��Ϊ��<%=saveok2%>Ԫ�� ���ڻ��ʱ����Ķ�㼸�Ǽ��֣��Ա����Ǹ��õ�ȷ�����Ļ��͸����Ϊ������ 
                                  ��<br>
                                  <br>
                                  <br>
                                  �ڻ������½ϵͳ�ύ�������ȷ�ϣ�������ȷ�Ϻ��24Сʱ�ڿ�ͨ���ķ����ٴθ�л�� <br>
                                  <br>
                                  <br>
                                  <br>
                                  <br>
                                  <br>
                                </p></td>
                            </tr>
                            <tr bgcolor=#FFFFFF>
                              <td height=22 bgcolor="#FFFFFF">&nbsp;</td>
                            </tr>
                          </table>                          <br>
                          <br>
                          <table width="100%" border="0" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td><br> <br> <br> <br> </td>
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
<%
else
response.Write("��������æ,���Ժ�����")
end if
%>
</body>
</html>