<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<!--#include file="css.asp"-->
<%on error resume next
%>
<%
if session("y_it_uid")="" then
response.redirect "error.asp?err=014&zurl=usermodview.asp"
response.end
end if
%>
<%
Set rs=Server.CreateObject("ADODB.RecordSet") 
sql="select * from user where uid='"&session("y_it_uid")&"'"
rs.open sql,conn,1,1
%>
<html>
<head>
<title><%=Application("y_it_itname")%>-�鿴����</title>
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
                                    <td valign="bottom" background="images/host-2.gif">�� �� �� ��</td>
                                  </tr>
                                </table>
                                <table width="75%" border="0" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td height="6"><img src="1.gif" width="1" height="1"></td>
                                  </tr>
                                </table>
                                
                                <table width="498" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                                  <tr bgcolor="#f5f5f5"> 
                                    <td height="38" colspan="2">�������ڴ˷�����޸����ĸ������ϣ�������������������룬�����´ε�½ʱʹ�������� 
                                      </td>
                                  </tr>
                                  <tr> 
                                    <td width="166" height="25" bgcolor="#f5f5f5">�û��ʺţ�</td>
                                    <td width="198" bgcolor="#FFFFFF"><%=rs("uid")%></td>
                                  </tr>
                                  <tr> 
                                    <td height="25" bgcolor="#f5f5f5">�����ʼ���</td>
                                    <td bgcolor="#FFFFFF"><%=rs("email")%> 
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td height="25" bgcolor="#f5f5f5">Ԥ���</td>
                                    <td bgcolor="#FFFFFF"><%=rs("prepay")%> RMB
									</td>
                                  </tr>
                                  <tr> 
                                    <td height="25" bgcolor="#f5f5f5">�Ѹ��</td>
                                    <td bgcolor="#FFFFFF"><%=rs("sumjifen")%> RMB
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td height="25" bgcolor="#f5f5f5">�Ա�</td>
                                    <td bgcolor="#FFFFFF"><%=rs("sex")%></td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">���䣺</td>
                                    <td bgcolor="#FFFFFF"><%=rs("age")%> 
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">��ҵ/��λ�����ģ���</td>
                                    <td bgcolor="#FFFFFF"><%=rs("namez")%> 
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">��ҵ/��λ��Ӣ�ģ���</td>
                                    <td bgcolor="#FFFFFF"><%=rs("namee")%> 
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">��ϵ�ˣ����ģ���</td>
                                    <td bgcolor="#FFFFFF"><%=rs("contactz")%> 
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">��ϵ�ˣ�Ӣ�ģ���</td>
                                    <td bgcolor="#FFFFFF"><%=rs("contacte")%> 
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">���ң����ģ���</td>
                                    <td bgcolor="#FFFFFF"><%=rs("govz")%> 
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5"> ���ң�Ӣ�ģ���</td>
                                    <td bgcolor="#FFFFFF"><%=rs("gove")%> 
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5"> ʡ�ݣ����ģ���</td>
                                    <td bgcolor="#FFFFFF"><%=rs("shengz")%> 
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">ʡ�ݣ�Ӣ�ģ���</td>
                                    <td bgcolor="#FFFFFF"><%=rs("shenge")%> 
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">���У����ģ���</td>
                                    <td bgcolor="#FFFFFF"><%=rs("cityz")%> 
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">���У�Ӣ�ģ���</td>
                                    <td bgcolor="#FFFFFF"><%=rs("citye")%> 
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">��ַ�����ģ���</td>
                                    <td bgcolor="#FFFFFF"><%=rs("dizhiz")%> 
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">��ַ��Ӣ�ģ���</td>
                                    <td bgcolor="#FFFFFF"><%=rs("dizhie")%> 
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">�������룺</td>
                                    <td bgcolor="#FFFFFF"><%=rs("postnumber")%> 
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">��ϵ�绰-1��</td>
                                    <td bgcolor="#FFFFFF"><%=rs("tel")%> 
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">��ϵ�绰-2��</td>
                                    <td bgcolor="#FFFFFF"><%=rs("tel2")%> 
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">OICQ/ICQ/MSN��</td>
                                    <td bgcolor="#FFFFFF"><%=rs("oicq")%> 
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">���棺</td>
                                    <td bgcolor="#FFFFFF"><%=rs("fax")%> 
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">&nbsp;</td>
                                    <td bgcolor="#FFFFFF">&nbsp;</td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#f5f5f5">&nbsp;</td>
                                    <td bgcolor="#FFFFFF">&nbsp;</td>
                                  </tr>
                                  <tr bgcolor="#FFFFFF"> 
                                    <td colspan="2">&nbsp;</td>
                                  </tr>
                                  <tr bgcolor="#f5f5f5"> 
                                    <td height="38" colspan="2"><div align="center"><a href="usermod.asp">����޸�>>></a></div></td>
                                  </tr>
                                </table></td>
                            </tr>
<%
rs.close
set rs=nothing
%>                            <tr> 
                              <td width="70%">&nbsp; </td>
                            </tr>
                          </table>
                          
                        </TD>
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
