<!--#include file="dbm_adminconn.asp" -->
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="dbm_css.css" type=text/css rel=stylesheet>
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
                    <TD height="40"><strong><font color="#FF0000">����������ƺ�̨����</font></strong></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
<%
If session("user_adminlevel")=1 Then
%>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#f5f5f5">
                <TBODY>
                  <TR> 
                    <TD height="25"><font color="#FF0000"><strong>����Ա����</strong></font></TD>
                  </TR>
                  <TR> 
                    <TD><a href="dbm_user_edit.asp?user_id=<%=session("user_id")%>" target="mainFrame">���������޸�</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="dbm_user_add.asp" target="mainFrame">����Ա ���</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="dbm_user_view.asp" target="mainFrame">����Ա �鿴�޸�</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="dbm_user_guest.asp" target="mainFrame">����Ա �����Ȩ��</a></TD>
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
                    <TD><a href="dbm_guest_add.asp" target="mainFrame">�ͻ� ���</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="dbm_guest_view.asp" target="mainFrame">�ͻ� �鿴�޸�</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5">
                <TBODY>
                  <TR>
                    <TD height="25"><strong><font color="#FF0000">��Ʒ</font></strong></TD>
                  </TR>
                  <TR> 
                    <TD width=118 height="25"><a href="dbm_subs_add.asp" target="mainFrame">��Ʒ���</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="dbm_subs_view.asp" target="mainFrame">��Ʒ�鿴�޸�</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
<%
sql="select * from dbm_message" 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
zrecordcount=rs.recordcount
rs.close
set rs=nothing
%>					
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5">
                <TBODY>
                  <TR>
                    <TD height="25"><strong><font color="#FF0000">����Ϣ (����<font color="#000000"><%=zrecordcount%></font>����Ϣ) </font></strong></TD>
                  </TR>
                  <TR> 
                    <TD width=118 height="25"><a href="dbm_message_add.asp" target="mainFrame">����Ϣ���</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="dbm_message_view.asp" target="mainFrame">����Ϣ�鿴�޸�</a></TD>
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
                    <TD width=118 height="25"><a href="dbm_note_add.asp" target="mainFrame">�������</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="dbm_note_view.asp" target="mainFrame">���Բ鿴�޸�</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5">
                <TBODY>
                  <TR>
                    <TD height="25"><strong><font color="#FF0000">�ղؼ�</font></strong></TD>
                  </TR>
                  <TR> 
                    <TD width=118 height="25"><a href="dbm_hildden_add.asp" target="mainFrame">�ղؼ����</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="dbm_hildden_view.asp" target="mainFrame">�ղؼв鿴�޸�</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
             <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5">
                <TBODY>
                  <TR> 
                    <TD width="118" height="35"><a href="folder_manager.asp" target="_blank"><strong><font color="#FF0000">�ļ�ϵͳ����</font></strong></a></TD>
                  </TR>
                </TBODY>
              </TABLE>
<%
End If
%>
<%
If session("user_adminlevel")=2 Then
%>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#f5f5f5">
                <TBODY>
                  <TR> 
                    <TD height="25"><font color="#FF0000"><strong>����Ա����</strong></font></TD>
                  </TR>
                  <TR> 
                    <TD><a href="dbm_user_edit.asp?user_id=<%=session("user_id")%>" target="mainFrame">���������޸�</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
<%
sql="select * from dbm_message where user_id="&session("user_id")
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
zrecordcount=rs.recordcount
rs.close
set rs=nothing
%>					
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5">
                <TBODY>
                  <TR>
                    <TD height="25"><strong><font color="#FF0000">����Ϣ (����<font color="#000000"><%=zrecordcount%></font>����Ϣ) </font></strong></TD>
                  </TR>
                  <TR> 
                    <TD width=118 height="25"><a href="dbm_message_add.asp" target="mainFrame">����Ϣ���</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="dbm_message_view.asp" target="mainFrame">����Ϣ�鿴�޸�</a></TD>
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
                    <TD width=118 height="25"><a href="dbm_note_add.asp" target="mainFrame">�������</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="dbm_note_view.asp" target="mainFrame">���Բ鿴�޸�</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5">
                <TBODY>
                  <TR> 
                    <TD width="118" height="35"><a href="folder_manager.asp" target="_blank"><strong><font color="#FF0000">�ļ�ϵͳ����</font></strong></a></TD>
                  </TR>
                </TBODY>
              </TABLE>
<%
End If
%>
<%
If session("user_adminlevel")=9 Then
%>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#f5f5f5">
                <TBODY>
                  <TR> 
                    <TD height="25"><font color="#FF0000"><strong>�ͻ�����</strong></font></TD>
                  </TR>
                  <TR> 
                    <TD><a href="dbm_guest_edit.asp?guest_id=<%=session("user_id")%>" target="mainFrame">���������޸�</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
<%
sql="select * from dbm_message where guest_id="&session("user_id")
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
zrecordcount=rs.recordcount
rs.close
set rs=nothing
%>					
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5">
                <TBODY>
                  <TR>
                    <TD height="25"><strong><font color="#FF0000">����Ϣ (����<font color="#000000"><%=zrecordcount%></font>����Ϣ) </font></strong></TD>
                  </TR>
                  <TR> 
                    <TD width=118 height="25"><a href="dbm_message_add.asp" target="mainFrame">����Ϣ���</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="dbm_message_view.asp" target="mainFrame">����Ϣ�鿴�޸�</a></TD>
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
                    <TD width=118 height="25"><a href="dbm_note_add.asp" target="mainFrame">�������</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="dbm_note_view.asp" target="mainFrame">���Բ鿴�޸�</a></TD>
                  </TR>
                </TBODY>
              </TABLE><br>
                            <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5">
                <TBODY>
                  <TR> 
                    <TD width="118" height="35"><a href="folder_manager.asp" target="_blank"><strong><font color="#FF0000">�ļ�ϵͳ����</font></strong></a></TD>
                  </TR>
                </TBODY>
              </TABLE>
<%
End If
%>
            </TD>
          </TR>
        </TBODY>
      </TABLE> </td>
        <td></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td height="50" align="center" valign="middle">
          <FORM action="dbm_login.asp" method="post" id="form_logout" name="form_logout" target="_parent">
            <INPUT type=submit value="logout" id="option" name="option" onClick="return confirm('ȷ���˳����رմ˴���!')">
          </FORM>
	</td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>