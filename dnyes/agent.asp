<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<!--#include file="css.asp"-->
<%on error resume next%>
<%
id=request("id")
sql="select * from myfaq where id="&id
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
newstype=rs("newstype")
subject=rs("subject")
message=rs("message")
message=replace(message,chr(13),"<br>")
idate=rs("idate")
rs.close
set rs=nothing
%>
<html>
<head>
<title><%=Application("y_it_itname")%>-<%=newstype%></title>
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
                <td valign="top"> 
                  <!--#include file="support.asp" -->
                </td>
                <td width="2" valign="top" background="images/left-2-right.gif"><img src="images/left-2-right.gif" width="2" height="5"></td>
              </tr>
            </table></td>
          <td width="7">&nbsp;</td>
          <td width="513" valign="top"><table width="100%" border="0" cellspacing="2" cellpadding="2">
              <tr> 
                <td width="3%"><font color="#003366">��</font></td>
                <td width="92%">��Ϊ���ſƼ����������</td>
                <td width="5%">&nbsp;</td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td width="25">��</td>
                      <td width="745"><font color="#0000FF">1�� ȫ�߲�Ʒ����С���ⶼ����</font><br>
                        ���ſƼ��Ĳ�Ʒ����80��Ʒ�֣���С��60��Ԫ���������Ԫ����Ϊ�����Ϳ�����Կͻ�û�й��ǣ����۴�С���������Գɽ������ض������ʡ�������Ȼ�����ܶࡣ<br> 
                      </td>
                    </tr>
                    <tr> 
                      <td>��</td>
                      <td><font color="#0000FF">2�� ������󣬻ر�Ѹ��</font><br>
                        ���ſƼ��Ĳ�Ʒ��������ߣ��ܶ��Ʒ����ߴ�35%������Ȼ����һ��׬һ����ʵʵ���ڵĽ��ˡ�<br> </td>
                    </tr>
                    <tr> 
                      <td>��</td>
                      <td><font color="#0000FF">3�� ȫ�����ʵ����֤</font><br>
                        ���ſƼ��ļ���ר��Ϊ�ͻ��ṩȫ���ķ��񣬽ڼ��ա����졢����Ҳ�����⣬�������С��˾���ֲ��㣬�豸�����ľ��档������վҵ����Խ��Խ�࣬�������ι�Ӧ�̷��񲻵�λ��ʧȥ�г������˿�Ҫ����������ľ���ʧ̫���ˡ�<br> 
                      </td>
                    </tr>
                    <tr> 
                      <td>��</td>
                      <td><font color="#0000FF">4�� �̳����ƣ�����г�������</font><br>
                        ���ſƼ���˾�Ǹ��г������͹�˾�����ſƼ����г��Ĳ���һֱ����ǿ�������������������ɡ����ԣ���λ��Ϊ�����Ϳ�����Ծ������ֺ���η�塣���ǵ����ƾ��Ǵ�ҵ����ƣ����ǳɳ��������һ����ɳ���</td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                    <tr> 
                      <td colspan="2">&nbsp;</td>
                    </tr>
                  </table></td>
                <td>&nbsp;</td>
              </tr>
            </table>
            <table width="100%" border="0" cellspacing="2" cellpadding="2">
              <tr> 
                <td width="3%"><font color="#003366">��</font></td>
                <td width="92%">���������</td>
                <td width="5%">&nbsp;</td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
                <td><TABLE border=0 cellPadding=2 cellSpacing=0 width="100%">
                    <TBODY>
                      <TR> 
                        <TD width=25>��</TD>
                        <TD width="737">������˴��������</TD>
                      </TR>
                      <TR> 
                        <TD width=25>&nbsp;</TD>
                        <TD>&nbsp;</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>�� </TD>
                        <TD>���Զ����е��������εĸ��ˡ�</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>�� </TD>
                        <TD>���Զ���Ϊ������չ���û����ݺϷ���Ч�ķ���Ʊ�ݡ�</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>�� </TD>
                        <TD>����Ϊ�û��ṩ��Ҫ�ļ�����������ѯ����</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>�� </TD>
                        <TD>Ӧ��ӵ�й̶��ķ�������</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>�� </TD>
                        <TD>���бȽϷḻ�Ļ������缼���������ҵ������</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>�� </TD>
                        <TD>���б���������ͨѶ��������Ҫ���豸��</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>�� </TD>
                        <TD>������Ҫ�ύ�������֤��ӡ����</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>�� </TD>
                        <TD>Ը�⽻�ɸ��˴���Ԥ����500ԪRMB�Ա�֤ҵ���˳����չ��</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>�� </TD>
                        <TD>Ը��������ſƼ��������ˡ��ල��ָ����</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>�� </TD>
                        <TD>�������ſƼ���"������ϵ�����淶"��</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>&nbsp;</TD>
                        <TD>&nbsp;</TD>
                      </TR>
                      <TR> 
                        <TD>��</TD>
                        <TD>������ҵ���������</TD>
                      </TR>
                      <TR> 
                        <TD>&nbsp;</TD>
                        <TD>&nbsp;</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>�� </TD>
                        <TD>���Զ����е��������ε���ҵ��λ��</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>�� </TD>
                        <TD>���Զ���Ϊ������չ���û����ݺϷ���Ч�ķ���Ʊ�ݡ�</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>�� </TD>
                        <TD>����Ϊ�û��ṩ��Ҫ�ļ�����������ѯ����</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>�� </TD>
                        <TD>Ӧ��ӵ�й̶��İ칫�ص㡣</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>�� </TD>
                        <TD>���бȽϷḻ�Ļ������缼���������ҵ������</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>�� </TD>
                        <TD>���б���������ͨѶ��������Ҫ���豸��</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>�� </TD>
                        <TD>��ҵ��Ҫ�ύ��ҵӪҵִ�ո�ӡ����</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>�� </TD>
                        <TD>�ڱ��ص���ҵ��ӵ��һ����֪���ȡ�</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>�� </TD>
                        <TD>�ڱ���ӵ�����Ƶ�����������������</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>�� </TD>
                        <TD>�������õĿͻ����������뾭�顣</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>�� </TD>
                        <TD>Ը�⽻�����ſƼ�����Ԥ����500ԪRMB�Ա�֤ҵ���˳����չ��</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>�� </TD>
                        <TD>Ը��������ſƼ��������ˡ��ල��ָ����</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>�� </TD>
                        <TD>�������ſƼ�"������ϵ�����淶"��</TD>
                      </TR>
                      <TR> 
                        <TD align=middle>�� </TD>
                        <TD>�����ſƼ������̵Ķ��������ų���ĳЩ���⹱�׵Ĵ�����ִ�����Կ��˵�</TD>
                      </TR>
                      <TR> 
                        <TD align=middle vAlign=top>&nbsp;</TD>
                        <TD>��׼��������ſƼ��ķ�չ���г��ƹ��������⹱�׻������Ҫ֧�ֵĴ�</TD>
                      </TR>
                      <TR> 
                        <TD align=middle vAlign=top>&nbsp;</TD>
                        <TD>���̣������ſƼ�ȷ�Ϻ󣬿������ṩ��Ӧ���Ŵ���֧�֡�</TD>
                      </TR>
                    </TBODY>
                  </TABLE></td>
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td><font color="#003366">��</font></td>
                <td>������������</td>
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td width="25">��</td>
                      <td width="745">�˽�Ҫ��ע���Ϊ�û� </td>
                    </tr>
                    <tr> 
                      <td>��</td>
                      <td>ϵͳ�Զ���֤Ϊ��ʽ�û�</td>
                    </tr>
                    <tr> 
                      <td>��</td>
                      <td><a href="mydocument/agent002.doc"><font color="#FF0000">���ش����ͬ</font></a>��ǩ�ֲ��ʼĵ��ҹ�˾</td>
                    </tr>
                    <tr> 
                      <td>��</td>
                      <td>��Ԥ�������ָ���ʻ�</td>
                    </tr>
                    <tr> 
                      <td>��</td>
                      <td>����ļ������ɨ������˾</td>
                    </tr>
                    <tr> 
                      <td>��</td>
                      <td>����������������ʽ�������ߴ��û�����Ϊ����</td>
                    </tr>
                  </table>
                </td>
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td><font color="#003366">��</font></td>
                <td>������˱�׼</td>
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
                <td><P>���ſƼ���������Ҫ����������˴����̺ͻ��������̣� ������������£�<BR>
                    1. ͭ�ƴ�����<BR>
                    2. ���ƴ�����<BR>
                    3. ���ƴ�����<BR>
                  </P>
                  <P><BR>
                    �¼���ı�׼</P>
                  <TABLE border=0 cellPadding=0 cellSpacing=1 
                  width="85%">
                    <TBODY>
                      <TR> 
                        <TD width="35%" height=20>������</TD>
                        <TD width="65%" height=20>��������</TD>
                      </TR>
                      <TR> 
                        <TD height=20 bgcolor="#ffffff">ͭ�ƴ�����</TD>
                        <TD height=20 bgcolor="#ffffff">Ԥ�տ��500Ԫ�����</TD>
                      </TR>
                      <TR> 
                        <TD height=20 bgcolor="#ffffff">���ƴ�����</TD>
                        <TD height=20 bgcolor="#ffffff">Ԥ�տ��3000Ԫ�����</TD>
                      </TR>
                      <TR> 
                        <TD height=20 bgcolor="#ffffff">���ƴ�����</TD>
                        <TD height=20 bgcolor="#ffffff">Ԥ�տ��10000Ԫ�����</TD>
                      </TR>
                    </TBODY>
                  </TABLE>
                  <br>
                  <br>
                  <a href="mydocument/agent002.doc"><font color="#FF0000">������ش����ͬ</font></a> 
                </td>
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
            </table> </td>
          <td width="7" background="images/right-7x5.gif">&nbsp;</td>
        </tr>
      </table></td>
    <td width="10" background="images/dnyes-bg-left-and-right.gif"><img src="images/dnyes-bg-left-and-right.gif" width="10" height="1"></td>
  </tr>
</table>
<!--#include file="copyright.asp"-->
</body>
</html>
