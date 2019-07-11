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
                    <TD height="40"><strong><font color="#FF0000">新筑建筑设计后台管理</font></strong></TD>
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
                    <TD height="25"><font color="#FF0000"><strong>管理员资料</strong></font></TD>
                  </TR>
                  <TR> 
                    <TD><a href="dbm_user_edit.asp?user_id=<%=session("user_id")%>" target="mainFrame">自我资料修改</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="dbm_user_add.asp" target="mainFrame">管理员 添加</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="dbm_user_view.asp" target="mainFrame">管理员 查看修改</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="dbm_user_guest.asp" target="mainFrame">管理员 管理的权限</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#CCCCCC" bgcolor="#f5f5f5">
                <TBODY>
                  <TR> 
                    <TD height="25"><font color="#FF0000"><strong>客户资料</strong></font></TD>
                  </TR>
                  <TR> 
                    <TD><a href="dbm_guest_add.asp" target="mainFrame">客户 添加</a></TD>
                  </TR>
                  <TR> 
                    <TD><a href="dbm_guest_view.asp" target="mainFrame">客户 查看修改</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5">
                <TBODY>
                  <TR>
                    <TD height="25"><strong><font color="#FF0000">作品</font></strong></TD>
                  </TR>
                  <TR> 
                    <TD width=118 height="25"><a href="dbm_subs_add.asp" target="mainFrame">作品添加</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="dbm_subs_view.asp" target="mainFrame">作品查看修改</a></TD>
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
                    <TD height="25"><strong><font color="#FF0000">短信息 (共有<font color="#000000"><%=zrecordcount%></font>条信息) </font></strong></TD>
                  </TR>
                  <TR> 
                    <TD width=118 height="25"><a href="dbm_message_add.asp" target="mainFrame">短信息添加</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="dbm_message_view.asp" target="mainFrame">短信息查看修改</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5">
                <TBODY>
                  <TR>
                    <TD height="25"><strong><font color="#FF0000">留言</font></strong></TD>
                  </TR>
                  <TR> 
                    <TD width=118 height="25"><a href="dbm_note_add.asp" target="mainFrame">留言添加</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="dbm_note_view.asp" target="mainFrame">留言查看修改</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5">
                <TBODY>
                  <TR>
                    <TD height="25"><strong><font color="#FF0000">收藏夹</font></strong></TD>
                  </TR>
                  <TR> 
                    <TD width=118 height="25"><a href="dbm_hildden_add.asp" target="mainFrame">收藏夹添加</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="dbm_hildden_view.asp" target="mainFrame">收藏夹查看修改</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
             <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5">
                <TBODY>
                  <TR> 
                    <TD width="118" height="35"><a href="folder_manager.asp" target="_blank"><strong><font color="#FF0000">文件系统管理</font></strong></a></TD>
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
                    <TD height="25"><font color="#FF0000"><strong>管理员资料</strong></font></TD>
                  </TR>
                  <TR> 
                    <TD><a href="dbm_user_edit.asp?user_id=<%=session("user_id")%>" target="mainFrame">自我资料修改</a></TD>
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
                    <TD height="25"><strong><font color="#FF0000">短信息 (共有<font color="#000000"><%=zrecordcount%></font>条信息) </font></strong></TD>
                  </TR>
                  <TR> 
                    <TD width=118 height="25"><a href="dbm_message_add.asp" target="mainFrame">短信息添加</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="dbm_message_view.asp" target="mainFrame">短信息查看修改</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5">
                <TBODY>
                  <TR>
                    <TD height="25"><strong><font color="#FF0000">留言</font></strong></TD>
                  </TR>
                  <TR> 
                    <TD width=118 height="25"><a href="dbm_note_add.asp" target="mainFrame">留言添加</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="dbm_note_view.asp" target="mainFrame">留言查看修改</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5">
                <TBODY>
                  <TR> 
                    <TD width="118" height="35"><a href="folder_manager.asp" target="_blank"><strong><font color="#FF0000">文件系统管理</font></strong></a></TD>
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
                    <TD height="25"><font color="#FF0000"><strong>客户资料</strong></font></TD>
                  </TR>
                  <TR> 
                    <TD><a href="dbm_guest_edit.asp?guest_id=<%=session("user_id")%>" target="mainFrame">自我资料修改</a></TD>
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
                    <TD height="25"><strong><font color="#FF0000">短信息 (共有<font color="#000000"><%=zrecordcount%></font>条信息) </font></strong></TD>
                  </TR>
                  <TR> 
                    <TD width=118 height="25"><a href="dbm_message_add.asp" target="mainFrame">短信息添加</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="dbm_message_view.asp" target="mainFrame">短信息查看修改</a></TD>
                  </TR>
                </TBODY>
              </TABLE>
              <br>
              <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5">
                <TBODY>
                  <TR>
                    <TD height="25"><strong><font color="#FF0000">留言</font></strong></TD>
                  </TR>
                  <TR> 
                    <TD width=118 height="25"><a href="dbm_note_add.asp" target="mainFrame">留言添加</a></TD>
                  </TR>
                  <TR> 
                    <TD height="25"><a href="dbm_note_view.asp" target="mainFrame">留言查看修改</a></TD>
                  </TR>
                </TBODY>
              </TABLE><br>
                            <TABLE width=100% border=1 cellpadding="5" cellspacing="0" bordercolor="#cccccc" bgcolor="#f5f5f5">
                <TBODY>
                  <TR> 
                    <TD width="118" height="35"><a href="folder_manager.asp" target="_blank"><strong><font color="#FF0000">文件系统管理</font></strong></a></TD>
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
            <INPUT type=submit value="logout" id="option" name="option" onClick="return confirm('确认退出并关闭此窗口!')">
          </FORM>
	</td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>