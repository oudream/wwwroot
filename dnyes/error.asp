<!--#include file="mymefaq5411jkjkh/favorible/showme.asp"-->
<%on error resume next%>
<%
errid=request("err")
if errid="" then err="001"
%>
<html>
<head>
<title><%=Application("y_it_itname")%>-���߶���</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<link rel="stylesheet" href="css.css" type="text/css">
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<!--#include file="top.asp" -->
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
          <td width="229" valign="top" background="images/left-228x5.gif"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="228" height="28" background="images/left-2.gif"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="10%" height="18" valign="bottom">&nbsp; </td>
                      <td width="8%" valign="bottom"><img src="images/gogo.gif" width="6" height="15"></td>
                      <td width="82%" valign="bottom"><font color="#FFFFFF"><b>�� 
                        �� ֪ ʶ</b></font></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td> <TABLE width=100% border=0 cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
                    <TBODY>
                      <TR> 
                        <TD height="2"
					vAlign=top 
                style="BORDER-left: #000000 1px solid"><img src="1.gif" width="1" height="1"></TD>
                      </TR>
                    </TBODY>
                  </TABLE>
                  <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td height="100" valign="top"> 
                        <!--#include file="base-index.asp" -->
                      </td>
                    </tr>
                  </table>
                  <TABLE width=100% border=0 cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
                    <TBODY>
                      <TR> 
                        <TD height="1"
				vAlign=top 
                style="BORDER-left: #000000 1px solid"><img src="1.gif" width="1" height="1"></TD>
                      </TR>
                    </TBODY>
                  </TABLE></td>
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
            <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td>
                  <!--#include file="support.asp" -->
                </td>
              </tr>
            </table> </td>
          <td width="7">&nbsp;</td>
          <td width="513" valign="top"><table width="100%" align="center">
              <tr align="center"> 
                <td align="center"> <br>
                  === <b>������ʾ</b> ===<br>
                  <br>
                  <br>
                </td>
              </tr>
              <tr> 
                <td> 
                  <table width="400" border="0" cellspacing="0" cellpadding="0" align="center">
                    <tr> 
                      <td width="50" valign="top"><br>
                        <br> 
                        <br> </td>
                      <td valign="top">���Ĳ������������⣬������ϸ�Ķ�������ʾ�� <br>
                        <br>
                        <br>
                        <%if errid="001" then%>
                        �����Ǵӱ����������ύ�ı���Ϣ������<br>
                        <br>
                        ������������������򱾲������Է����������ƻ���<br>
                        <br>
                        ���������һ�����ߵ�Ȩ������<br>
                        <br>
                        <br>
                        <br>
                        �����κ����������ţ�<a href="mailto:<%=Application("y_it_fromemail")%>"><%=Application("y_it_fromemail")%></a> 
                        <br> 
                        <%end if%>
                        <%if errid="002" then%>
<script language="JavaScript">
alert("�Բ���!    �����ʺŻ������벻��ȷ\n\n����ȷ��д�����ʺ�����");
history.go( -1 );
</script>
                        <%end if%>
                        <%if errid="003" then%>
<script language="JavaScript">
alert("�Բ���!    ����ѯ����/���ݲ�����");
history.go( -1 );
</script>
                        <%end if%>
                        <%if errid="004" then%>
                        ���û����ѱ�ע��<br>
                        <br>
                        ����ѡ���������û���<br>
                        <br>
                        ��л��ʹ�ñ�ϵͳ<br>
                        <br>
                        �����κ����������ţ�<a href="mailto:<%=Application("y_it_fromemail")%>"><%=Application("y_it_fromemail")%></a> 
                        <br> 
                        <%end if%>
                        <%if errid="005" then%>
                        �Բ����һ��������<br>
                        <br>
                        ϵͳ�в������ڴ��ʺ�<br>
                        <br>
                        ��л��ʹ�ñ�ϵͳ<br>
                        <br>
                        �����κ����������ţ�<a href="mailto:<%=Application("y_it_fromemail")%>"><%=Application("y_it_fromemail")%></a> 
                        <br>
                        <%end if%>
                        <%if errid="006" then%>
                        <font color=red>����δ��¼</font>����Ӵ���ڵ�¼����<a href="userregprotocal.asp">���ע��</a><br>
                        <br> <form action="login.asp" method="post">
                          �����ʺ�:
                          <input type="text" name="uid" size="16" maxlength="20" class="form">
                          &nbsp;
                          <input type="submit" name="Submit" class=botton value="��¼">
                          <input type="hidden" name="url" value="add.asp?add=<%=request("add")%>">
                          <br>
                          ��������:
                          <input type="password" name="pwd" size="16" maxlength="20" class="form">
                          &nbsp;
                          <input class=botton type="submit" name="Submit2"  value="ȡ��">
                        </form>
                        <br>
                        �����κ����������ţ�<a href="mailto:<%=Application("y_it_fromemail")%>"><%=Application("y_it_fromemail")%></a> 
                        <br>
                        <%end if%>
                        <%if errid="007" then%>
                        �Բ�����û��������������Ѿ�ȷ�Ϲ���������˽�<br>
						��鿴���Ķ��������¹���<br>
                        <br>
						����ʱ���������<a href="servicefollow.asp">��������</a><br><br>
                        ��л�������ǵ�֧��<br>
                        <br>
                        <br>
                        <br>
                        �����κ����������ţ�<a href="mailto:<%=Application("y_it_fromemail")%>"><%=Application("y_it_fromemail")%></a> 
                        <br> 
                        <%end if%>
                        <%if errid="008" then%>
                        �Բ������Ĺ����ܽ����Ͳ�����Ԥ����<br>
                        <br>
						����������ʽ�����л�������ǵ�֧�֣�
                        <br>
                        <br>
                        �����κ����������ţ�<a href="mailto:<%=Application("y_it_fromemail")%>"><%=Application("y_it_fromemail")%></a> 
                        <br> 
                        <%end if%>
                        <%if errid="009" then%>
                        �Բ������Ķ����Ѿ������߹��ˣ��벻Ҫ�ظ�����<br>
                        <br>
                        ��л�������ǵ�֧�֣�<br>
                        <br>
                        �����κ����������ţ�<a href="mailto:<%=Application("y_it_fromemail")%>"><%=Application("y_it_fromemail")%></a>����QQ��30013002,30013004 
                        �� <br> 
                        <%end if%>
                        <%if errid="010" then%>
                        �Բ��𣬲��޴˶�������ȷ��<br>
                        <br>
                        ������ȷʹ�ñ�ϵͳ<br>
                        <br>
                        ��л�������ǵ�֧�֣�<br>
                        <br>
                        �����κ����������ţ�<a href="mailto:<%=Application("y_it_fromemail")%>"><%=Application("y_it_fromemail")%></a>����QQ��30013002,30013004 
                        �� <br> 
                        <%end if%>
                        <%if errid="011" then%>
                        �Բ���<font color=red>����δ��¼</font>����Ӵ���ڵ�¼����<a href="reg.asp">���ע��</a><br>
                        <br> <form action="login.asp" method="post">
                          �����ʺ�:
                          <input type="text" name="uid" size="16" maxlength="20" class="form">
                          &nbsp;
                          <input type="submit" name="Submit" class=botton value="��¼">
                          <input type="hidden" name="url" value="index.asp">
                          <br>
                          ��������:
                          <input type="password" name="pwd" size="16" maxlength="20" class="form">
                          &nbsp;
                          <input type="submit" name="Submit2" class=botton value="ȡ��">
                        </form>
                        <br>
                        �����κ����������ţ�<a href="mailto:<%=Application("y_it_fromemail")%>"><%=Application("y_it_fromemail")%></a> 
                        <br> 
                        <%end if%>
                        <br>
                        <%if errid="012" then%>
                        �Բ���<font color=red>����δ��¼</font>����Ӵ���ڵ�¼����<a href="userregprotocal.asp">���ע��</a><br>
                        <br>
                        <br> <form action="login.asp" method="post">
                          �����ʺ�:
                          <input type="text" name="uid" size="16" maxlength="20" class="form">
                          &nbsp;
                          <input type="submit" name="Submit" class=botton value="��¼">
                          <input type="hidden" name="url" value="<%=session("domainurl")%>">
                          <br>
                          ��������:
                          <input type="password" name="pwd" size="16" maxlength="20" class="form">
                          &nbsp;
                          <input type="submit" name="Submit2" class=botton value="ȡ��">
                        </form>
                        <br>
                        <br>
                        <br>
                        �����κ����������ţ�<a href="mailto:<%=Application("y_it_fromemail")%>"><%=Application("y_it_fromemail")%></a> 
                        <br> 
                        <%end if%>
						<br>
                        <%if errid="013" then%>
                        �Բ���<font color="red">���ѵ�¼</font><br>
                        <br>
                        ������ע���û�,������ע���ȡ������,����ע��<br>
                        <br>
                        ��л�������ǵ�֧�֣�<br>
                        <br>
                        �����κ����������ţ�<a href="mailto:<%=Application("y_it_fromemail")%>"><%=Application("y_it_fromemail")%></a> 
                        <br> 
                        <%end if%>
                        <br>
                        <%if errid="014" then%>
                        �Բ���<font color=red>����δ��¼</font>����Ӵ���ڵ�¼����<a href="userregprotocal.asp">���ע��</a><br>
                        <br>
                        <br> <form action="login.asp" method="post">
                          �����ʺ�:
                          <input type="text" name="uid" size="16" maxlength="20" class="form">
                          &nbsp;
                          <input type="submit" name="Submit" class=botton value="��¼">
                          <input type="hidden" name="url" value="<%=request("zurl")%>">
                          <br>
                          ��������:
                          <input type="password" name="pwd" size="16" maxlength="20" class="form">
                          &nbsp;
                          <input type="submit" name="Submit2" class=botton value="ȡ��">
                        </form>
                        <br>
                        <br>
                        <br>
                        �����κ����������ţ�<a href="mailto:<%=Application("y_it_fromemail")%>"><%=Application("y_it_fromemail")%></a> 
                        <br> 
                        <%end if%>
                        <br> 
						</td>
                    </tr>
                  </table> </td>
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
