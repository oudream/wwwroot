

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="195" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td>&nbsp;</td>
	    
    <td align="center"><TABLE height=100% width=100% border=0>
        <TBODY>
          <TR> 
            <TD vAlign=top align=left> <TABLE width=100% height="30" border=0>
                <TBODY>
                  <TR> 
                    <TD height="40"><strong><font color="#FF0000">SOSOON</font></strong></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#f5f5f5">
                <TBODY>
                  <TR> 
                    <TD height="25"><font color="#FF0000"><strong>�û�����</strong></font></TD>
                  </TR>
                  <TR> 
                    <TD><a href="user_edit.asp" target="mainFrame">���������޸�</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="user_view.asp" target="mainFrame">�鿴�����û�</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#f5f5f5">
                <TBODY>
                  <TR> 
                    <TD height="25"><font color="#FF0000"><strong>�ͻ�����</strong></font></TD>
                  </TR>
                  <TR> 
                    <TD><a href="guest_add.asp" target="mainFrame">�ͻ����</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="guest_view.asp" target="mainFrame">�ͻ��鿴�޸�</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5">
                <TBODY>
                  <TR>
                    <TD height="25"><strong><font color="#FF0000">����</font></strong></TD>
                  </TR>
                  <TR> 
                    <TD width=169 height="25"><a href="sort_add.asp" target="mainFrame">�������</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="sort_view.asp" target="mainFrame">����鿴�޸�</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="sort_brand.asp" target="mainFrame">��������Ʒ�ƵĲ鿴�޸�</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5">
                <TBODY>
                  <TR>
                    <TD height="25"><strong><font color="#FF0000">Ʒ��</font></strong></TD>
                  </TR>
                  <TR> 
                    <TD width=118 height="25"><a href="brand_add.asp" target="mainFrame">Ʒ�����</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="brand_view.asp" target="mainFrame">Ʒ�Ʋ鿴�޸�</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <%
if session("user_adminlevel")=1 then
%>
              <%
end if
%>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5">
                <TBODY>
                  <TR>
                    <TD height="25"><strong><font color="#FF0000">��Ʒ</font></strong></TD>
                  </TR>
                  <TR> 
                    <TD width=118 height="25"><a href="subs_add.asp" target="mainFrame">��Ʒ���</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="subs_view.asp" target="mainFrame">��Ʒ�鿴�޸�</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
			  <br>
             <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5">
                <TBODY>
                  <TR> 
                    <TD width="118" height="35"><a href="subs_send_price.asp" target="mainFrame"><strong><font color="#FF0000">��������</font></strong></a></TD>
                  </TR>
                  <TR> 
                    <TD width="118" height="35"><a href="sosoon_document.asp" target="mainFrame"><strong><font color="#FF0000">���ŵ���</font></strong></a></TD>
                  </TR>
                  <TR> 
                    <TD width="118" height="35"><a href="sosoon_note_view.asp" target="mainFrame"><strong><font color="#FF0000">���ű���</font></strong></a></TD>
                  </TR>
                  <TR> 
                    <TD width="118" height="35"><a href="message_view.asp" target="mainFrame"><strong><font color="#FF0000">��������</font></strong></a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              
            </TD>
          </TR>
        </TBODY>
      </TABLE> </td>
        <td></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td height="50" align="center" valign="middle">
          <FORM action="login.asp" method="post" id="form_logout" name="form_logout" target="_parent">
            <INPUT type=submit value="logout" id="option" name="option">
          </FORM>
	</td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>